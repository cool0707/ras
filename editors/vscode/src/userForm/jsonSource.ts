// A small JSON parser that keeps the source range of every value, so that the
// Problems panel can reveal the field a validation issue refers to, and that
// rejects duplicate keys (JSON.parse silently keeps the last one).
//
// Part of the UserForm Designer ported from xlflow
// (https://github.com/harumiWeb/xlflow, MIT License, Copyright (c) 2026 harumiWeb).

export type JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue };

export interface JsonNode {
	start: number;
	end: number;
	/** Range of the key of a member value (objects only). */
	keyStart?: number;
	members?: Map<string, JsonNode>;
	elements?: JsonNode[];
}

export class JsonSyntaxError extends Error {
	constructor(
		message: string,
		readonly offset: number,
	) {
		super(message);
	}
}

export function parseJson(text: string): { value: JsonValue; node: JsonNode } {
	let pos = 0;
	const fail = (message: string): never => {
		throw new JsonSyntaxError(message, pos);
	};
	const space = () => {
		while (pos < text.length && " \t\r\n﻿".includes(text[pos])) pos++;
	};
	const value = (): { value: JsonValue; node: JsonNode } => {
		space();
		const start = pos;
		const ch = text[pos];
		if (ch === "{") {
			pos++;
			const result: { [key: string]: JsonValue } = {};
			const members = new Map<string, JsonNode>();
			space();
			if (text[pos] === "}") {
				pos++;
				return { value: result, node: { start, end: pos, members } };
			}
			for (;;) {
				space();
				const keyStart = pos;
				if (text[pos] !== '"') fail("オブジェクトのキーが必要です");
				const key = string();
				if (members.has(key)) {
					pos = keyStart;
					fail(`キー "${key}" が重複しています`);
				}
				space();
				if (text[pos] !== ":") fail("':' が必要です");
				pos++;
				const member = value();
				member.node.keyStart = keyStart;
				Object.defineProperty(result, key, { value: member.value, enumerable: true, writable: true, configurable: true });
				members.set(key, member.node);
				space();
				if (text[pos] === ",") {
					pos++;
					continue;
				}
				if (text[pos] === "}") {
					pos++;
					return { value: result, node: { start, end: pos, members } };
				}
				fail("',' または '}' が必要です");
			}
		}
		if (ch === "[") {
			pos++;
			const result: JsonValue[] = [];
			const elements: JsonNode[] = [];
			space();
			if (text[pos] === "]") {
				pos++;
				return { value: result, node: { start, end: pos, elements } };
			}
			for (;;) {
				const element = value();
				result.push(element.value);
				elements.push(element.node);
				space();
				if (text[pos] === ",") {
					pos++;
					continue;
				}
				if (text[pos] === "]") {
					pos++;
					return { value: result, node: { start, end: pos, elements } };
				}
				fail("',' または ']' が必要です");
			}
		}
		if (ch === '"') return { value: string(), node: { start, end: pos } };
		const literal = /^(?:true|false|null|-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?)/.exec(text.slice(pos, pos + 64));
		if (!literal) return fail("値が必要です");
		pos += literal[0].length;
		const word = literal[0];
		const parsed: JsonValue = word === "true" ? true : word === "false" ? false : word === "null" ? null : Number(word);
		return { value: parsed, node: { start, end: pos } };
	};
	const string = (): string => {
		const start = pos;
		pos++;
		while (pos < text.length && text[pos] !== '"') {
			if (text[pos] === "\\") pos++;
			else if (text.charCodeAt(pos) < 0x20) fail("文字列に制御文字があります");
			pos++;
		}
		if (pos >= text.length) fail("文字列が閉じていません");
		pos++;
		try {
			return JSON.parse(text.slice(start, pos)) as string;
		} catch {
			pos = start;
			return fail("文字列のエスケープが正しくありません");
		}
	};
	const result = value();
	space();
	if (pos < text.length) fail("余分な文字があります");
	return result;
}

/** Locates `form.font.size` / `controls[3].tabs[0].name` style paths. */
export function nodeAtPath(root: JsonNode, path: string): JsonNode | undefined {
	let node: JsonNode | undefined = root;
	let last: JsonNode | undefined = root;
	for (const part of path.split(/\.|(?=\[)/).filter(Boolean)) {
		if (!node) break;
		const index = /^\[(\d+)\]$/.exec(part);
		node = index ? node.elements?.[Number(index[1])] : node.members?.get(part);
		if (node) last = node;
	}
	return node ?? last;
}

export function offsetToPosition(text: string, offset: number): { line: number; character: number } {
	let line = 0;
	let lineStart = 0;
	for (let i = 0; i < offset && i < text.length; i++) {
		if (text[i] === "\n") {
			line++;
			lineStart = i + 1;
		}
	}
	return { line, character: offset - lineStart };
}

/** Serialises with the indentation and line ending already used by `previous`. */
export function formatJson(value: unknown, previous = ""): string {
	const indentMatch = /\n([ \t]+)"/.exec(previous);
	const indent = indentMatch ? indentMatch[1] : "  ";
	const newline = previous.includes("\r\n") ? "\r\n" : "\n";
	return JSON.stringify(value, null, indent).replace(/\n/g, newline) + newline;
}

/** One replacement covering the changed middle part (common prefix/suffix kept). */
export function minimalEdit(before: string, after: string): { start: number; end: number; text: string } | undefined {
	if (before === after) return undefined;
	let start = 0;
	const limit = Math.min(before.length, after.length);
	while (start < limit && before.charCodeAt(start) === after.charCodeAt(start)) start++;
	let endBefore = before.length;
	let endAfter = after.length;
	while (endBefore > start && endAfter > start && before.charCodeAt(endBefore - 1) === after.charCodeAt(endAfter - 1)) {
		endBefore--;
		endAfter--;
	}
	return { start, end: endBefore, text: after.slice(start, endAfter) };
}

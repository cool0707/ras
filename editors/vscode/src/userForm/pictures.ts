// Image control pictures: validation of BMP / JPEG files (the same rules as
// the compiler), project asset paths, and previews for the Designer canvas.
// This module does not depend on VS Code.

import * as fs from "node:fs/promises";
import * as path from "node:path";
import type { ParsedSpec, ValidationIssue } from "./spec";
import { isObject } from "./spec";

export const MAX_PICTURE_BYTES = 16 << 20;
export const MAX_PICTURE_PIXELS = 16_000_000;
/** Previews of one form stop being attached beyond this many source bytes. */
export const MAX_PREVIEW_TOTAL_BYTES = 32 << 20;
/** Project-root-relative folder that receives pictures picked outside the project. */
export const ASSET_FOLDER = "forms/assets";

/** A picture that the compiler would reject; the message is shown to the user. */
export class PictureError extends Error {}

export interface PictureInfo {
	format: "bmp" | "jpeg";
	width: number;
	height: number;
}

const sizeMessage = "画像ファイルは 1 バイト以上 16 MiB 以下でなければなりません。";

function checkDimensions(width: number, height: number) {
	if (width * height > MAX_PICTURE_PIXELS)
		throw new PictureError(`画像の大きさ（${width} × ${height}）が上限の 16,000,000 ピクセルを超えています。`);
}

function inspectBmp(data: Uint8Array, view: DataView): PictureInfo {
	if (data.length < 54) throw new PictureError("BMP のヘッダーが途中で切れています。");
	if (view.getUint32(14, true) !== 40) throw new PictureError("BITMAPINFOHEADER（40 バイトのヘッダー）を持つ BMP だけを使えます。");
	const width = view.getInt32(18, true);
	const signedHeight = view.getInt32(22, true);
	if (width <= 0 || signedHeight === 0 || signedHeight === -0x80000000) throw new PictureError("BMP の幅と高さが正しくありません。");
	const height = Math.abs(signedHeight);
	checkDimensions(width, height);
	if (view.getUint16(26, true) !== 1) throw new PictureError("BMP のプレーン数は 1 でなければなりません。");
	const bits = view.getUint16(28, true);
	if (bits !== 24 && bits !== 32) throw new PictureError(`24 ビットまたは 32 ビットの BMP だけを使えます（この画像は ${bits} ビットです）。`);
	if (view.getUint32(30, true) !== 0) throw new PictureError("無圧縮（BI_RGB）の BMP だけを使えます。");
	const pixelOffset = view.getUint32(10, true);
	const pixelBytes = Math.floor((width * bits + 31) / 32) * 4 * height;
	if (pixelOffset < 54 || pixelOffset + pixelBytes > data.length) throw new PictureError("BMP のピクセル データがファイルに収まっていません。");
	// The sizes the header declares may be 0 (unknown) but must not contradict the pixel data.
	const declaredSize = view.getUint32(2, true);
	if (declaredSize !== 0 && (declaredSize > data.length || declaredSize < pixelOffset + pixelBytes))
		throw new PictureError("BMP のヘッダーにあるファイル サイズがピクセル データと合いません。");
	const imageSize = view.getUint32(34, true);
	if (imageSize !== 0 && imageSize < pixelBytes) throw new PictureError("BMP のヘッダーにある画像サイズがピクセル データより小さくなっています。");
	return { format: "bmp", width, height };
}

/** Walks the marker segments up to the start of the scan; the image data itself is not decoded. */
function inspectJpeg(data: Uint8Array, view: DataView): PictureInfo {
	const broken = (what: string) => new PictureError(`JPEG として読めません（${what}）。`);
	let at = 2;
	let frame: { width: number; height: number } | undefined;
	for (;;) {
		// Fill bytes (0xFF) may precede a marker.
		while (at + 1 < data.length && data[at] === 0xff && data[at + 1] === 0xff) at++;
		if (at + 4 > data.length || data[at] !== 0xff) throw broken("マーカーが見つかりません");
		const marker = data[at + 1];
		const length = view.getUint16(at + 2, false);
		if (length < 2 || at + 2 + length > data.length) throw broken("セグメントが途中で切れています");
		if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
			// A frame header (SOFn): precision, height, width.
			if (length < 8) throw broken("フレーム ヘッダーが短すぎます");
			frame = { height: view.getUint16(at + 5, false), width: view.getUint16(at + 7, false) };
		} else if (marker === 0xda) break;
		else if (marker === 0xd8 || marker === 0xd9) throw broken("マーカーの並びが正しくありません");
		at += 2 + length;
	}
	if (!frame) throw broken("フレーム ヘッダーがありません");
	// The scan must be closed by an end-of-image marker.
	let closed = false;
	for (let i = at; i + 1 < data.length && !closed; i++) closed = data[i] === 0xff && data[i + 1] === 0xd9;
	if (!closed) throw broken("終了マーカーがありません");
	if (frame.width === 0 || frame.height === 0) throw new PictureError("JPEG の幅と高さが正しくありません。");
	checkDimensions(frame.width, frame.height);
	return { format: "jpeg", ...frame };
}

/** Checks that `data` is a BMP or JPEG file the compiler can embed. */
export function inspectPicture(data: Uint8Array): PictureInfo {
	if (data.length < 1 || data.length > MAX_PICTURE_BYTES) throw new PictureError(sizeMessage);
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	if (data.length >= 2 && data[0] === 0x42 && data[1] === 0x4d) return inspectBmp(data, view);
	if (data.length >= 2 && data[0] === 0xff && data[1] === 0xd8) return inspectJpeg(data, view);
	throw new PictureError("BMP または JPEG の画像ファイルを指定してください。");
}

export function pictureDataUri(data: Uint8Array, format: PictureInfo["format"]): string {
	return `data:image/${format};base64,${Buffer.from(data.buffer, data.byteOffset, data.byteLength).toString("base64")}`;
}

/** Reads a picture file without ever loading more than the size limit. */
export async function readPictureFile(file: string): Promise<{ data: Uint8Array; info: PictureInfo }> {
	let handle: fs.FileHandle;
	try {
		handle = await fs.open(file, "r");
	} catch {
		throw new PictureError("画像ファイルが見つからないか、開けません。");
	}
	try {
		const stat = await handle.stat();
		if (!stat.isFile()) throw new PictureError("画像ファイルが見つからないか、開けません。");
		if (stat.size < 1 || stat.size > MAX_PICTURE_BYTES) throw new PictureError(sizeMessage);
		// One spare byte reveals a file that grew after stat().
		const buffer = Buffer.alloc(stat.size + 1);
		let length = 0;
		for (;;) {
			const { bytesRead } = await handle.read(buffer, length, buffer.length - length, length);
			if (!bytesRead) break;
			length += bytesRead;
			if (length === buffer.length) break;
		}
		if (length !== stat.size) throw new PictureError("画像ファイルが読み取り中に変更されました。もう一度選んでください。");
		const data = buffer.subarray(0, length);
		return { data, info: inspectPicture(data) };
	} catch (error) {
		if (error instanceof PictureError) throw error;
		throw new PictureError("画像ファイルを読み取れません。");
	} finally {
		await handle.close();
	}
}

// ---------------------------------------------------------------------------
// Project paths
// ---------------------------------------------------------------------------

const exists = (file: string) =>
	fs.stat(file).then(
		(stat) => stat.isFile(),
		() => false,
	);

/**
 * The project root of a FormSpec: the nearest ancestor folder that holds
 * `ras.toml`, or the parent of the spec's folder (`forms`) when there is none.
 */
export async function findProjectRoot(specFile: string): Promise<string> {
	const forms = path.dirname(path.resolve(specFile));
	for (let folder = forms; ; ) {
		if (await exists(path.join(folder, "ras.toml"))) return folder;
		const parent = path.dirname(folder);
		if (parent === folder) return path.dirname(forms);
		folder = parent;
	}
}

/** `file` as a root-relative path with `/` separators, or undefined when it is not inside `root`. */
export function projectRelativePath(root: string, file: string): string | undefined {
	const relative = path.relative(path.resolve(root), path.resolve(file));
	if (!relative || path.isAbsolute(relative) || relative === ".." || relative.startsWith(`..${path.sep}`)) return undefined;
	return relative.split(path.sep).join("/");
}

/**
 * The file a `picture.path` names, or undefined when the compiler would not
 * take the path: it must be relative, without a drive or `..`.
 */
export function resolvePicturePath(root: string, picturePath: string): string | undefined {
	const normalized = picturePath.replace(/\\/g, "/");
	if (!picturePath.trim() || normalized.startsWith("/") || normalized.includes(":")) return undefined;
	const parts = normalized.split("/").filter((part) => part !== "" && part !== ".");
	if (!parts.length || parts.includes("..")) return undefined;
	return path.join(path.resolve(root), ...parts);
}

export interface PictureAssetPlan {
	/** The value for `picture.path`. */
	relativePath: string;
	/** Where to copy the picked file first; absent when an existing project file is referenced. */
	copyTo?: string;
}

async function sameContent(file: string, data: Uint8Array): Promise<boolean> {
	try {
		const stat = await fs.stat(file);
		if (!stat.isFile() || stat.size !== data.length) return false;
		return Buffer.compare(await fs.readFile(file), data) === 0;
	} catch {
		return false;
	}
}

/** Refuses a plan whose path the compiler would not read back as the same file (e.g. a `:` in the name). */
function usable(root: string, plan: PictureAssetPlan): PictureAssetPlan {
	if (plan.relativePath.includes("\\") || resolvePicturePath(root, plan.relativePath) === undefined)
		throw new PictureError(`"${plan.relativePath}" は画像のパスとして使えません。ファイル名やフォルダー名に : や \\ を含めないでください。`);
	return plan;
}

/**
 * Decides how a picked file is referenced. A file inside the project is used
 * in place. Any other file goes to `forms/assets/<name>`; when a different
 * file already has that name the copy is named `<stem>-2.<ext>`, `-3`, ...,
 * and an identical file that is already there is reused. Nothing is written.
 */
export async function planPictureAsset(root: string, source: string, data: Uint8Array): Promise<PictureAssetPlan> {
	// Links are followed first: a link that leads out of the project is not a
	// project file, and the project may itself be opened through a link.
	const [realRoot, realSource] = await Promise.all([root, source].map((p) => fs.realpath(p).catch(() => path.resolve(p))));
	const real = projectRelativePath(realRoot, realSource);
	if (real !== undefined) return usable(root, { relativePath: projectRelativePath(root, source) ?? real });
	const assets = path.join(root, ...ASSET_FOLDER.split("/"));
	// Names are compared without case so that the result is the same on every file system.
	const taken = new Map<string, string>();
	for (const name of await fs.readdir(assets).catch(() => [] as string[])) taken.set(name.toLowerCase(), name);
	const { name: stem, ext } = path.parse(path.basename(source));
	for (let n = 1; n <= 9999; n++) {
		const candidate = n === 1 ? `${stem}${ext}` : `${stem}-${n}${ext}`;
		const existing = taken.get(candidate.toLowerCase());
		if (existing === undefined) return usable(root, { relativePath: `${ASSET_FOLDER}/${candidate}`, copyTo: path.join(assets, candidate) });
		if (await sameContent(path.join(assets, existing), data)) return usable(root, { relativePath: `${ASSET_FOLDER}/${existing}` });
	}
	throw new PictureError(`${ASSET_FOLDER} に同じ名前のファイルが多すぎます。ファイル名を変えてください。`);
}

/** Performs the copy of a plan. Never overwrites: fails when the target appeared meanwhile. */
export async function writePictureAsset(plan: PictureAssetPlan, data: Uint8Array): Promise<void> {
	if (!plan.copyTo) return;
	await fs.mkdir(path.dirname(plan.copyTo), { recursive: true });
	try {
		await fs.writeFile(plan.copyTo, data, { flag: "wx" });
	} catch (error) {
		// A partly written copy is removed; a file that someone else created is kept.
		if ((error as NodeJS.ErrnoException).code !== "EEXIST") await fs.rm(plan.copyTo, { force: true }).catch(() => undefined);
		throw error;
	}
}

// ---------------------------------------------------------------------------
// Canvas previews
// ---------------------------------------------------------------------------

export interface PicturePreview {
	dataUri?: string;
	/** Source size of the picture, for the per-form preview budget. */
	bytes?: number;
	/** Why there is no preview. */
	problem?: string;
}

/** Previews by absolute file path. A file is read again only after it changed on disk. */
export class PicturePreviews {
	private readonly entries = new Map<string, { stamp: string; preview: PicturePreview }>();

	get(file: string): PicturePreview | undefined {
		return this.entries.get(file)?.preview;
	}

	/**
	 * Loads one file of the project `root` unless the cached preview is current.
	 * Returns true when the preview changed.
	 */
	async load(file: string, root: string): Promise<boolean> {
		let stamp = "missing";
		try {
			const stat = await fs.stat(file);
			if (stat.isFile()) stamp = `${stat.mtimeMs}:${stat.size}`;
		} catch {
			// Reported as a missing file below.
		}
		const cached = this.entries.get(file);
		if (cached?.stamp === stamp) return false;
		let preview: PicturePreview;
		try {
			// Like the compiler, do not follow a link out of the project.
			const [realRoot, realFile] = await Promise.all([root, file].map((p) => fs.realpath(p).catch(() => undefined)));
			if (realRoot && realFile && projectRelativePath(realRoot, realFile) === undefined) throw new PictureError("リンクの先がプロジェクトの外にあります。");
			const { data, info } = await readPictureFile(file);
			preview = { dataUri: pictureDataUri(data, info.format), bytes: data.length };
		} catch (error) {
			preview = { problem: error instanceof Error ? error.message : String(error) };
		}
		this.entries.set(file, { stamp, preview });
		return cached?.preview.dataUri !== preview.dataUri || cached?.preview.problem !== preview.problem;
	}

	/** Brings the cache in line with the files a form uses now. Returns true when any preview changed. */
	async refresh(files: Iterable<string>, root: string): Promise<boolean> {
		const wanted = new Set(files);
		for (const file of [...this.entries.keys()]) if (!wanted.has(file)) this.entries.delete(file);
		let changed = false;
		for (const file of wanted) if (await this.load(file, root)) changed = true;
		return changed;
	}
}

/** The `picture.path` of each Image control that has one, by index into `spec.records`. */
function picturePaths(spec: ParsedSpec): { index: number; field: string; picturePath: string }[] {
	const result: { index: number; field: string; picturePath: string }[] = [];
	spec.records.forEach((record, index) => {
		const picture = record.raw.picture;
		if (record.control.type.toLowerCase() !== "image" || !isObject(picture)) return;
		if (typeof picture.path !== "string" || !picture.path.trim() || picture.remove === true) return;
		result.push({ index, field: `${record.path}.picture.path`, picturePath: picture.path });
	});
	return result;
}

/** Absolute paths of the picture files a form refers to. */
export function pictureFiles(spec: ParsedSpec, root: string): string[] {
	const files = new Set<string>();
	for (const { picturePath } of picturePaths(spec)) {
		const file = resolvePicturePath(root, picturePath);
		if (file) files.add(file);
	}
	return [...files];
}

/**
 * What the loaded previews mean for a form: a data URI per control (aligned
 * with `spec.records`) and a warning for each picture that cannot be shown.
 */
export function picturePreviewState(spec: ParsedSpec, root: string, previews: PicturePreviews): { dataUris: (string | undefined)[]; issues: ValidationIssue[] } {
	const dataUris: (string | undefined)[] = spec.records.map(() => undefined);
	const issues: ValidationIssue[] = [];
	const counted = new Set<string>();
	let total = 0;
	const warn = (field: string, message: string, suggestion: string) => issues.push({ code: "UFP001", severity: "warning", message, field, suggestion });
	for (const { index, field, picturePath } of picturePaths(spec)) {
		const file = resolvePicturePath(root, picturePath);
		if (!file) {
			warn(field, `${field} "${picturePath}" はプロジェクト ルートからの相対パスではありません（.. や絶対パスは使えません）。`, "ras.toml のあるフォルダーからの相対パスを / 区切りで指定してください。");
			continue;
		}
		const preview = previews.get(file);
		if (preview?.problem) warn(field, `${field} "${picturePath}" を画像として使えません: ${preview.problem}`, "24 / 32 ビットの無圧縮 BMP か JPEG のファイルを指定してください。");
		if (!preview?.dataUri) continue;
		if (!counted.has(file)) {
			if (total + (preview.bytes ?? 0) > MAX_PREVIEW_TOTAL_BYTES) continue;
			counted.add(file);
			total += preview.bytes ?? 0;
		}
		dataUris[index] = preview.dataUri;
	}
	return { dataUris, issues };
}

// Pictures of controls and forms (`picture`, `mouseIcon`): validation of picture files (the
// checks of the compiler that need no decoder), project asset paths, and previews for the
// Designer canvas. This module does not depend on VS Code.

import * as fs from "node:fs/promises";
import * as path from "node:path";
import type { ParsedSpec, ValidationIssue } from "./spec";
import { isObject } from "./spec";
import { controlProperties } from "./contract";

export const MAX_PICTURE_BYTES = 16 << 20;
export const MAX_PICTURE_PIXELS = 16_000_000;
/** Previews of one form stop being attached beyond this many source bytes. */
export const MAX_PREVIEW_TOTAL_BYTES = 32 << 20;
/** Project-root-relative folder that receives pictures picked outside the project. */
export const ASSET_FOLDER = "forms/assets";

/** A picture that the compiler would reject; the message is shown to the user. */
export class PictureError extends Error {}

export type PictureFormat = "bmp" | "jpeg" | "png" | "gif" | "webp" | "tiff" | "ico" | "cur";

export interface PictureInfo {
	format: PictureFormat;
	/** Size in pixels, where the file header gives it without decoding. */
	width?: number;
	height?: number;
	/**
	 * False for a picture that is embedded as it is (a 24 / 32 bit uncompressed
	 * BMP, a JPEG); true for one the compiler converts to a 24-bit BMP when it
	 * writes the form.
	 */
	converted: boolean;
}

export const FORMAT_LIST = "BMP / JPEG / PNG / GIF / WebP / TIFF / ICO";
/** File name extensions offered by the open dialog. */
export const PICTURE_EXTENSIONS = ["bmp", "jpg", "jpeg", "png", "gif", "webp", "tif", "tiff", "ico"];
/** File name extensions offered for a `mouseIcon`: icons and cursors, stored as they are. */
export const ICON_EXTENSIONS = ["ico", "cur"];

const sizeMessage = "画像ファイルは 1 バイト以上 16 MiB 以下でなければなりません。";

function checkDimensions(width: number, height: number) {
	if (width * height > MAX_PICTURE_PIXELS)
		throw new PictureError(`画像の大きさ（${width} × ${height}）が上限の 16,000,000 ピクセルを超えています。`);
}

/** The check of a BMP that MS Forms can store unchanged. */
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
	return { format: "bmp", width, height, converted: false };
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
	return { format: "jpeg", ...frame, converted: false };
}

/** Size of the 24-bit BMP a picture of this many pixels is converted to. */
export function convertedBmpBytes(width: number, height: number): number {
	return 54 + Math.floor((width * 3 + 3) / 4) * 4 * height;
}

/** A picture that is converted; its size is only known where the header is trivial to read. */
function converted(format: PictureFormat, width?: number, height?: number): PictureInfo {
	if (width === undefined || height === undefined) return { format, converted: true };
	if (width === 0 || height === 0) throw new PictureError("画像の幅と高さが正しくありません。");
	if (convertedBmpBytes(width, height) > MAX_PICTURE_BYTES)
		throw new PictureError(`画像が大きすぎます（${width} × ${height}）。24 ビットの BMP に変換すると 16 MiB を超えます。`);
	return { format, width, height, converted: true };
}

const startsWith = (data: Uint8Array, bytes: number[], at = 0) => data.length >= at + bytes.length && bytes.every((byte, index) => data[at + index] === byte);
const ascii = (text: string) => [...text].map((ch) => ch.charCodeAt(0));

/**
 * Checks a picture file as far as that is possible without decoding it:
 * the size, the format (by signature, never by file name) and, for the
 * formats whose header is trivial to read, the size of the converted bitmap.
 * Damaged image data is left for the compiler to report.
 */
export function inspectPicture(data: Uint8Array): PictureInfo {
	if (data.length < 1 || data.length > MAX_PICTURE_BYTES) throw new PictureError(sizeMessage);
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	if (startsWith(data, ascii("BM"))) {
		try {
			return inspectBmp(data, view);
		} catch {
			// 8-bit, RLE, V4 / V5 headers and the like are converted. The size is
			// taken from the header only when it is plainly there.
			const header = data.length >= 26 ? view.getUint32(14, true) : 0;
			const [width, height] = header === 12 ? [view.getUint16(18, true), view.getUint16(20, true)] : header >= 40 ? [view.getInt32(18, true), Math.abs(view.getInt32(22, true))] : [0, 0];
			return width > 0 && height > 0 ? converted("bmp", width, height) : converted("bmp");
		}
	}
	// A JPEG is always embedded as it is, so one that fails the check is an error.
	if (startsWith(data, [0xff, 0xd8])) return inspectJpeg(data, view);
	if (startsWith(data, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
		// The IHDR chunk comes first: width and height at offsets 16 and 20.
		return startsWith(data, ascii("IHDR"), 12) && data.length >= 24 ? converted("png", view.getUint32(16, false), view.getUint32(20, false)) : converted("png");
	}
	if (startsWith(data, ascii("GIF87a")) || startsWith(data, ascii("GIF89a"))) {
		// The logical screen descriptor follows the signature.
		return data.length >= 10 ? converted("gif", view.getUint16(6, true), view.getUint16(8, true)) : converted("gif");
	}
	if (startsWith(data, ascii("RIFF")) && startsWith(data, ascii("WEBP"), 8)) return converted("webp");
	if (startsWith(data, [0x49, 0x49, 0x2a, 0x00]) || startsWith(data, [0x4d, 0x4d, 0x00, 0x2a])) return converted("tiff");
	if (startsWith(data, [0x00, 0x00, 0x01, 0x00])) return converted("ico");
	throw new PictureError(`画像の形式を判別できません（${FORMAT_LIST} に対応しています）。`);
}

/**
 * Checks an icon or cursor file for a `mouseIcon` (stored as it is). A BMP or JPEG that
 * MS Forms holds unchanged is taken too; nothing is converted.
 */
export function inspectIcon(data: Uint8Array): PictureInfo {
	if (data.length < 1 || data.length > MAX_PICTURE_BYTES) throw new PictureError(sizeMessage);
	const icon = startsWith(data, [0x00, 0x00, 0x01, 0x00]);
	if (icon || startsWith(data, [0x00, 0x00, 0x02, 0x00])) {
		const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
		const count = data.length >= 6 ? view.getUint16(4, true) : 0;
		if (count === 0 || data.length < 6 + count * 16) throw new PictureError("アイコン（ICO / CUR）として読めません（画像の一覧が途中で切れています）。");
		for (let entry = 0; entry < count; entry++) {
			const size = view.getUint32(6 + entry * 16 + 8, true);
			const offset = view.getUint32(6 + entry * 16 + 12, true);
			if (size === 0 || offset < 6 + count * 16 || offset + size > data.length) throw new PictureError("アイコン（ICO / CUR）として読めません（画像のデータがファイルの外を指しています）。");
		}
		const side = (b: number) => (b === 0 ? 256 : b);
		return { format: icon ? "ico" : "cur", width: side(data[6]), height: side(data[7]), converted: false };
	}
	const info = (() => {
		try {
			return inspectPicture(data);
		} catch {
			return undefined;
		}
	})();
	if (info && !info.converted) return info;
	throw new PictureError("MouseIcon にはアイコン（.ico）かカーソル（.cur）のファイルを指定してください。");
}

const mimeTypes: Record<PictureFormat, string> = {
	bmp: "image/bmp",
	jpeg: "image/jpeg",
	png: "image/png",
	gif: "image/gif",
	webp: "image/webp",
	tiff: "image/tiff",
	ico: "image/x-icon",
	cur: "image/x-icon",
};

export function pictureDataUri(data: Uint8Array, format: PictureFormat): string {
	return `data:${mimeTypes[format]};base64,${Buffer.from(data.buffer, data.byteOffset, data.byteLength).toString("base64")}`;
}

/** Whether the webview (Chromium) can draw the format; TIFF is the one it cannot. */
export const previewable = (format: PictureFormat) => format !== "tiff";

/** Reads a picture file (an icon or cursor for a `mouseIcon`) without ever loading more than the size limit. */
export async function readPictureFile(file: string, icon = false): Promise<{ data: Uint8Array; info: PictureInfo }> {
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
		return { data, info: icon ? inspectIcon(data) : inspectPicture(data) };
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
	/** Why there is no preview (the file is no picture). */
	problem?: string;
	/** Why the file cannot be a mouse icon. */
	iconProblem?: string;
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
			const { data } = await readPictureFile(file, true).catch(() => readPictureFile(file));
			const message = (error: unknown) => (error instanceof Error ? error.message : String(error));
			let iconProblem: string | undefined;
			try {
				inspectIcon(data);
			} catch (error) {
				iconProblem = message(error);
			}
			try {
				const info = inspectPicture(data);
				// A format the webview cannot draw has no preview; that is not a problem of the form.
				preview = previewable(info.format) ? { dataUri: pictureDataUri(data, info.format), bytes: data.length } : {};
			} catch (error) {
				preview = { problem: message(error) };
			}
			if (iconProblem) preview.iconProblem = iconProblem;
		} catch (error) {
			const problem = error instanceof Error ? error.message : String(error);
			preview = { problem, iconProblem: problem };
		}
		this.entries.set(file, { stamp, preview });
		return cached?.preview.dataUri !== preview.dataUri || cached?.preview.problem !== preview.problem || cached?.preview.iconProblem !== preview.iconProblem;
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

/** A picture action of a form: of a control (by index into `spec.records`) or of the form (-1). */
interface PictureUse {
	index: number;
	property: "picture" | "mouseIcon";
	field: string;
	picturePath: string;
}

/** The `picture.path` and `mouseIcon.path` of the form and of each control that has them. */
function picturePaths(spec: ParsedSpec): PictureUse[] {
	const result: PictureUse[] = [];
	const add = (owner: Record<string, unknown>, index: number, base: string) => {
		for (const property of ["picture", "mouseIcon"] as const) {
			const picture = owner[property];
			if (!isObject(picture) || typeof picture.path !== "string" || !picture.path.trim() || picture.remove === true) continue;
			result.push({ index, property, field: `${base}.${property}.path`, picturePath: picture.path });
		}
	};
	if (isObject(spec.form)) add(spec.form, -1, "form");
	spec.records.forEach((record, index) => {
		const properties = controlProperties(record.control.type);
		const owner = Object.fromEntries(Object.entries(record.raw).filter(([key]) => key in properties));
		add(owner, index, record.path);
	});
	return result;
}

/** Absolute paths of the picture files a form refers to (pictures and mouse icons). */
export function pictureFiles(spec: ParsedSpec, root: string): string[] {
	const files = new Set<string>();
	for (const { picturePath } of picturePaths(spec)) {
		const file = resolvePicturePath(root, picturePath);
		if (file) files.add(file);
	}
	return [...files];
}

/** The previews of a form: of each control's picture (aligned with `spec.records`) and of the form's. */
export interface PicturePreviewState {
	dataUris: (string | undefined)[];
	/** The form's background picture. */
	form?: string;
	issues: ValidationIssue[];
}

/**
 * What the loaded previews mean for a form: a data URI per picture, and a warning for each
 * picture or mouse icon that cannot be used.
 */
export function picturePreviewState(spec: ParsedSpec, root: string, previews: PicturePreviews): PicturePreviewState {
	const state: PicturePreviewState = { dataUris: spec.records.map(() => undefined), issues: [] };
	const counted = new Set<string>();
	let total = 0;
	const warn = (field: string, message: string, suggestion: string) => state.issues.push({ code: "UFP001", severity: "warning", message, field, suggestion });
	for (const { index, property, field, picturePath } of picturePaths(spec)) {
		const file = resolvePicturePath(root, picturePath);
		if (!file) {
			warn(field, `${field} "${picturePath}" はプロジェクト ルートからの相対パスではありません（.. や絶対パスは使えません）。`, "ras.toml のあるフォルダーからの相対パスを / 区切りで指定してください。");
			continue;
		}
		const preview = previews.get(file);
		const problem = property === "mouseIcon" ? preview?.iconProblem : preview?.problem;
		if (problem) warn(field, `${field} "${picturePath}" を${property === "mouseIcon" ? "マウス アイコン" : "画像"}として使えません: ${problem}`, property === "mouseIcon" ? "ICO / CUR のファイルを指定してください。" : `${FORMAT_LIST} のファイルを指定してください。`);
		// A mouse icon is not drawn on the canvas.
		if (property === "mouseIcon" || !preview?.dataUri) continue;
		if (!counted.has(file)) {
			if (total + (preview.bytes ?? 0) > MAX_PREVIEW_TOTAL_BYTES) continue;
			counted.add(file);
			total += preview.bytes ?? 0;
		}
		if (index < 0) state.form = preview.dataUri;
		else state.dataUris[index] = preview.dataUri;
	}
	return state;
}

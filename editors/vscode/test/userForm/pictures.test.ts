import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { after, describe, it } from "node:test";
import { formatJson } from "../../src/userForm/jsonSource";
import {
	convertedBmpBytes,
	findProjectRoot,
	ICON_EXTENSIONS,
	inspectIcon,
	inspectPicture,
	MAX_PICTURE_BYTES,
	PICTURE_EXTENSIONS,
	pictureDataUri,
	PictureError,
	pictureFiles,
	PicturePreviews,
	picturePreviewState,
	planPictureAsset,
	previewable,
	projectRelativePath,
	readPictureFile,
	resolvePicturePath,
	writePictureAsset,
} from "../../src/userForm/pictures";
import { newFormSpec, parseSpec, problems } from "../../src/userForm/spec";

/** An uncompressed BMP with a 40-byte header; `edit` can then break it. */
function bmp(width: number, height: number, bits = 24, edit?: (data: Buffer) => void): Buffer {
	const rowBytes = Math.floor((width * bits + 31) / 32) * 4;
	const data = Buffer.alloc(54 + rowBytes * Math.abs(height), 0x7f);
	data.fill(0, 0, 54);
	data.write("BM", 0, "latin1");
	data.writeUInt32LE(data.length, 2);
	data.writeUInt32LE(54, 10);
	data.writeUInt32LE(40, 14);
	data.writeInt32LE(width, 18);
	data.writeInt32LE(height, 22);
	data.writeUInt16LE(1, 26);
	data.writeUInt16LE(bits, 28);
	edit?.(data);
	return data;
}

const segment = (marker: number, payload: number[]) => [0xff, marker, (payload.length + 2) >> 8, (payload.length + 2) & 0xff, ...payload];
const frame = (marker: number, width: number, height: number) => segment(marker, [8, height >> 8, height & 0xff, width >> 8, width & 0xff, 1, 1, 0x11, 0]);
const app0 = segment(0xe0, [0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0]);
const scan = [...segment(0xda, [1, 1, 0, 0, 63, 0]), 0x12, 0x34];
const eoi = [0xff, 0xd9];
/** SOI and the given segments, without scan data or the end marker. */
const jpegHead = (...parts: number[][]) => Buffer.from([0xff, 0xd8, ...parts.flat()]);
const jpeg = (...parts: number[][]) => jpegHead(...parts, scan, eoi);

/** The signature and the IHDR chunk of a PNG; the image data is left out. */
function png(width: number, height: number): Buffer {
	const data = Buffer.alloc(33);
	Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]).copy(data);
	data.write("IHDR", 12, "latin1");
	data.writeUInt32BE(width, 16);
	data.writeUInt32BE(height, 20);
	data.set([8, 6, 0, 0, 0], 24);
	return data;
}

/** The signature and the logical screen descriptor of a GIF. */
function gif(version: string, width: number, height: number): Buffer {
	const data = Buffer.alloc(13);
	data.write(version, 0, "latin1");
	data.writeUInt16LE(width, 6);
	data.writeUInt16LE(height, 8);
	return data;
}

const rejects = (data: Uint8Array, pattern: RegExp) =>
	assert.throws(
		() => inspectPicture(data),
		(error: unknown) => error instanceof PictureError && pattern.test(error.message),
	);

describe("picture file validation", () => {
	it("accepts uncompressed 24-bit and 32-bit BMP files", () => {
		assert.deepEqual(inspectPicture(bmp(3, 2)), { format: "bmp", width: 3, height: 2, converted: false });
		assert.deepEqual(inspectPicture(bmp(5, -4, 32)), { format: "bmp", width: 5, height: 4, converted: false });
		// Bytes after the pixel array and a later pixel offset are fine.
		const padded = Buffer.concat([bmp(2, 2), Buffer.alloc(8)]);
		assert.equal(inspectPicture(padded).format, "bmp");
		assert.ok(pictureDataUri(bmp(1, 1), "bmp").startsWith("data:image/bmp;base64,Qk"));
	});

	it("accepts other BMP files as pictures to convert", () => {
		const toConvert = (data: Uint8Array, size: { width?: number; height?: number } = { width: 3, height: 2 }) =>
			assert.deepEqual(inspectPicture(data), { format: "bmp", ...size, converted: true });
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt16LE(8, 28))); // 8-bit
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt32LE(1, 30))); // RLE
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt32LE(124, 14))); // V5 header
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt16LE(0, 26)));
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt32LE(40, 10)));
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt32LE(58, 10)));
		toConvert(bmp(3, 2).subarray(0, 60));
		toConvert(bmp(3, 2).subarray(0, 40));
		// Sizes in the header may be 0; ones that contradict the pixel data make it a picture to convert.
		assert.equal(inspectPicture(bmp(3, 2, 24, (d) => d.writeUInt32LE(0, 2))).converted, false);
		assert.equal(inspectPicture(bmp(3, 2, 24, (d) => d.writeUInt32LE(24, 34))).converted, false);
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt32LE(d.length + 1, 2)));
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt32LE(60, 2)));
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt32LE(23, 34)));
		// An OS/2 header has 16-bit sizes.
		toConvert(bmp(3, 2, 24, (d) => (d.writeUInt32LE(12, 14), d.writeUInt16LE(7, 18), d.writeUInt16LE(5, 20))), { width: 7, height: 5 });
		// A header whose size cannot be trusted is left to the compiler.
		toConvert(bmp(3, 2, 24, (d) => d.writeInt32LE(0, 18)), {});
		toConvert(bmp(3, 2, 24, (d) => d.writeInt32LE(-3, 18)), {});
		toConvert(bmp(3, 2, 24, (d) => d.writeInt32LE(0, 22)), {});
		toConvert(bmp(3, 2, 24, (d) => d.writeUInt32LE(16, 14)), {});
		toConvert(bmp(3, 2).subarray(0, 20), {});
		toConvert(Buffer.from("BM"), {});
		// The converted bitmap must fit, whatever the bit depth of the source.
		rejects(bmp(1, 1, 24, (d) => (d.writeUInt16LE(8, 28), d.writeInt32LE(3000, 18), d.writeInt32LE(-3000, 22))), /3000 × 3000.*16 MiB/);
		assert.equal(inspectPicture(bmp(1, 1, 24, (d) => (d.writeUInt16LE(8, 28), d.writeInt32LE(2364, 18), d.writeInt32LE(2364, 22)))).width, 2364);
	});

	it("reads JPEG dimensions from the first frame header", () => {
		const table = segment(0xc4, [0, 1, 2, 3]);
		assert.deepEqual(inspectPicture(jpeg(app0, frame(0xc0, 24, 16))), { format: "jpeg", width: 24, height: 16, converted: false });
		// DHT (0xC4) is not a frame header; progressive frames (0xC2) count.
		assert.deepEqual(inspectPicture(jpeg(app0, table, frame(0xc2, 640, 480), table)), { format: "jpeg", width: 640, height: 480, converted: false });
		// Fill bytes before a marker are skipped.
		assert.deepEqual(inspectPicture(jpeg(app0, [0xff, 0xff], frame(0xc1, 300, 2))), { format: "jpeg", width: 300, height: 2, converted: false });
		assert.deepEqual(inspectPicture(jpeg(frame(0xc0, 4000, 4000))), { format: "jpeg", width: 4000, height: 4000, converted: false });
	});

	it("rejects JPEG files without usable dimensions", () => {
		rejects(jpeg(app0, frame(0xc0, 0, 16)), /幅と高さ/);
		rejects(jpeg(app0, frame(0xc0, 24, 0)), /幅と高さ/);
		rejects(jpeg(app0, frame(0xc0, 5000, 4000)), /16,000,000/);
		// No frame header before the scan.
		rejects(jpeg(app0), /フレーム ヘッダーがありません/);
		rejects(jpegHead(app0, scan, frame(0xc0, 24, 16), eoi), /フレーム ヘッダーがありません/);
		// The file must reach the scan and be closed by the end marker.
		rejects(jpegHead(app0, frame(0xc0, 24, 16)), /マーカーが見つかりません/);
		rejects(jpegHead(app0, frame(0xc0, 24, 16), scan), /終了マーカー/);
		rejects(jpegHead(app0, frame(0xc0, 24, 16), eoi, scan, eoi), /途中で切れています/);
		rejects(jpegHead(app0, frame(0xc0, 24, 16), [0xff, 0xd9, 0, 2], scan, eoi), /マーカーの並び/);
		rejects(jpegHead(app0, frame(0xc0, 24, 16).slice(0, 6)), /途中で切れています/);
		rejects(jpeg(app0, segment(0xc0, [8, 0, 16, 0])), /短すぎます/);
		rejects(jpeg(app0, [0x00, 0x01, 0x02, 0x03]), /マーカーが見つかりません/);
		rejects(jpeg([0xff, 0xe0, 0, 1]), /途中で切れています/);
	});

	it("detects the formats to convert by signature", () => {
		const signature = (...bytes: (number | string)[]) => Buffer.concat(bytes.map((b) => (typeof b === "string" ? Buffer.from(b, "latin1") : Buffer.from([b]))));
		assert.deepEqual(inspectPicture(png(48, 32)), { format: "png", width: 48, height: 32, converted: true });
		assert.deepEqual(inspectPicture(gif("GIF89a", 300, 200)), { format: "gif", width: 300, height: 200, converted: true });
		assert.deepEqual(inspectPicture(gif("GIF87a", 1, 1)), { format: "gif", width: 1, height: 1, converted: true });
		assert.deepEqual(inspectPicture(signature("RIFF", 0x24, 0, 0, 0, "WEBPVP8 ")), { format: "webp", converted: true });
		assert.deepEqual(inspectPicture(signature("II*", 0, 8, 0, 0, 0)), { format: "tiff", converted: true });
		assert.deepEqual(inspectPicture(signature("MM", 0, "*", 0, 0, 0, 8)), { format: "tiff", converted: true });
		assert.deepEqual(inspectPicture(signature(0, 0, 1, 0, 1, 0, 16, 16)), { format: "ico", converted: true });
		// A header too short to give the size is left to the compiler.
		assert.deepEqual(inspectPicture(png(48, 32).subarray(0, 20)), { format: "png", converted: true });
		assert.deepEqual(inspectPicture(Buffer.from("GIF89a")), { format: "gif", converted: true });
		// The file name plays no part: only these signatures are pictures.
		const unknown = /判別できません（BMP \/ JPEG \/ PNG \/ GIF \/ WebP \/ TIFF \/ ICO に対応しています）/;
		rejects(Buffer.from("hello"), unknown);
		rejects(Buffer.from("B"), unknown);
		rejects(signature("GIF88a", 1, 0, 1, 0), unknown);
		rejects(signature("RIFF", 0x24, 0, 0, 0, "WAVEfmt "), unknown);
		rejects(signature("II", 0, "*"), unknown);
		rejects(signature(0, 0, 2, 0, 1, 0), unknown);
		rejects(signature(0x89, "PNG\r\n"), unknown);
		rejects(signature("<svg xmlns='http://www.w3.org/2000/svg'/>"), unknown);
		// A JPEG is embedded as it is, so a broken one is not handed to the converter.
		rejects(signature(0xff, 0xd8, 0xff, 0xe0), /JPEG として読めません/);
	});

	it("limits pictures to convert by the size of the resulting 24-bit BMP", () => {
		assert.equal(convertedBmpBytes(1, 1), 58);
		assert.equal(convertedBmpBytes(3, 2), 78);
		assert.equal(convertedBmpBytes(4, 1), 66);
		// 2364 x 2364 pixels fit into 16 MiB as a 24-bit BMP, 2365 x 2365 do not.
		assert.ok(convertedBmpBytes(2364, 2364) <= MAX_PICTURE_BYTES && convertedBmpBytes(2365, 2365) > MAX_PICTURE_BYTES);
		assert.equal(inspectPicture(png(2364, 2364)).width, 2364);
		rejects(png(2365, 2365), /2365 × 2365.*16 MiB/);
		rejects(png(0x7fffffff, 0x7fffffff), /16 MiB/);
		assert.equal(inspectPicture(gif("GIF89a", 2364, 2364)).height, 2364);
		rejects(gif("GIF89a", 2365, 2365), /16 MiB/);
		rejects(gif("GIF87a", 65535, 65535), /16 MiB/);
		rejects(png(0, 32), /幅と高さ/);
		rejects(gif("GIF89a", 16, 0), /幅と高さ/);
		// Pictures embedded as they are keep the 16,000,000 pixel limit.
		assert.equal(inspectPicture(jpeg(frame(0xc0, 4000, 4000))).converted, false);
	});

	it("gives each format its MIME type and leaves TIFF without a preview", () => {
		const types = (["bmp", "jpeg", "png", "gif", "webp", "tiff", "ico"] as const).map((format) => pictureDataUri(Buffer.from("x"), format));
		assert.deepEqual(types, ["image/bmp", "image/jpeg", "image/png", "image/gif", "image/webp", "image/tiff", "image/x-icon"].map((type) => `data:${type};base64,eA==`));
		assert.deepEqual(["bmp", "jpeg", "png", "gif", "webp", "tiff", "ico"].filter((format) => !previewable(format as "bmp")), ["tiff"]);
		assert.deepEqual(PICTURE_EXTENSIONS, ["bmp", "jpg", "jpeg", "png", "gif", "webp", "tif", "tiff", "ico"]);
	});

	it("rejects sizes outside 1 byte to 16 MiB", () => {
		rejects(Buffer.alloc(0), /16 MiB/);
		const large = Buffer.alloc(MAX_PICTURE_BYTES + 1);
		large.write("BM", 0, "latin1");
		rejects(large, /16 MiB/);
	});
});

describe("picture assets", () => {
	const base = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), "ras-pictures-")));
	after(() => fs.rmSync(base, { recursive: true, force: true }));
	const write = (file: string, data: string | Uint8Array = "") => {
		fs.mkdirSync(path.dirname(file), { recursive: true });
		fs.writeFileSync(file, data);
		return file;
	};
	const project = (name: string) => {
		const root = path.join(base, name);
		write(path.join(root, "ras.toml"));
		write(path.join(root, "forms", "Main.form.json"), "{}");
		return root;
	};

	it("finds the project root from ras.toml, or falls back to the parent of forms", async () => {
		const root = project("root");
		assert.equal(await findProjectRoot(path.join(root, "forms", "Main.form.json")), root);
		// The nearest ras.toml wins.
		const inner = path.join(root, "packages", "inner");
		write(path.join(inner, "ras.toml"));
		assert.equal(await findProjectRoot(path.join(inner, "forms", "A.form.json")), inner);
		assert.equal(await findProjectRoot(path.join(root, "packages", "other", "forms", "A.form.json")), root);
		const loose = path.join(base, "loose", "forms", "A.form.json");
		assert.equal(await findProjectRoot(loose), path.join(base, "loose"));
	});

	it("turns files inside the project into / separated relative paths", () => {
		const root = path.join(base, "root");
		assert.equal(projectRelativePath(root, path.join(root, "forms", "assets", "logo.bmp")), "forms/assets/logo.bmp");
		assert.equal(projectRelativePath(root, path.join(root, "logo.jpg")), "logo.jpg");
		assert.equal(projectRelativePath(root, root), undefined);
		assert.equal(projectRelativePath(root, path.join(base, "elsewhere", "logo.bmp")), undefined);
		assert.equal(projectRelativePath(root, path.join(base, "root-2", "logo.bmp")), undefined);
		assert.equal(resolvePicturePath(root, "forms/assets/logo.bmp"), path.join(root, "forms", "assets", "logo.bmp"));
		// Like the compiler: backslashes are separators, and "." segments are ignored.
		assert.equal(resolvePicturePath(root, "forms\\assets\\logo.bmp"), path.join(root, "forms", "assets", "logo.bmp"));
		assert.equal(resolvePicturePath(root, "./forms//logo.bmp"), path.join(root, "forms", "logo.bmp"));
		for (const bad of ["", " ", "../logo.bmp", "forms/../logo.bmp", "forms\\..\\..\\logo.bmp", "/logo.bmp", "\\logo.bmp", "C:\\logo.bmp", "C:logo.bmp", "a:b.bmp", ".", "./"])
			assert.equal(resolvePicturePath(root, bad), undefined, bad);
	});

	it("references project files in place and copies other files to forms/assets", async () => {
		const root = project("copy");
		const assets = path.join(root, "forms", "assets");
		const inside = write(path.join(root, "images", "logo.bmp"), bmp(2, 2));
		assert.deepEqual(await planPictureAsset(root, inside, bmp(2, 2)), { relativePath: "images/logo.bmp" });

		const first = bmp(2, 2);
		const outside = write(path.join(base, "downloads", "logo.bmp"), first);
		const plan = await planPictureAsset(root, outside, first);
		assert.deepEqual(plan, { relativePath: "forms/assets/logo.bmp", copyTo: path.join(assets, "logo.bmp") });
		// Planning writes nothing.
		assert.equal(fs.existsSync(assets), false);
		await writePictureAsset(plan, first);
		assert.deepEqual(fs.readFileSync(path.join(assets, "logo.bmp")), first);

		// The same picture again reuses the copy.
		assert.deepEqual(await planPictureAsset(root, outside, first), { relativePath: "forms/assets/logo.bmp" });

		// A different picture of the same name gets a numbered name.
		const second = bmp(3, 3);
		const planSecond = await planPictureAsset(root, outside, second);
		assert.deepEqual(planSecond, { relativePath: "forms/assets/logo-2.bmp", copyTo: path.join(assets, "logo-2.bmp") });
		await writePictureAsset(planSecond, second);
		const third = bmp(4, 4);
		assert.equal((await planPictureAsset(root, outside, third)).relativePath, "forms/assets/logo-3.bmp");
		assert.deepEqual(await planPictureAsset(root, outside, second), { relativePath: "forms/assets/logo-2.bmp" });

		// Names that differ only in case count as the same name.
		const upper = write(path.join(base, "downloads", "LOGO.bmp"), third);
		assert.equal((await planPictureAsset(root, upper, third)).relativePath, "forms/assets/LOGO-3.bmp");
		assert.deepEqual(await planPictureAsset(root, upper, first), { relativePath: "forms/assets/logo.bmp" });
		assert.equal((await planPictureAsset(root, write(path.join(base, "downloads", "noext"), third), third)).relativePath, "forms/assets/noext");
	});

	it("copies a file that only a link places inside the project", async (t) => {
		const root = project("links");
		const data = bmp(2, 2);
		const target = write(path.join(base, "linked", "photo.bmp"), data);
		try {
			fs.symlinkSync(path.dirname(target), path.join(root, "shared"), "dir");
			fs.symlinkSync(root, path.join(base, "alias"), "dir");
		} catch {
			return t.skip("symbolic links are not available");
		}
		// The compiler does not follow a link out of the project, so the file is copied.
		assert.equal((await planPictureAsset(root, path.join(root, "shared", "photo.bmp"), data)).relativePath, "forms/assets/photo.bmp");
		const previews = new PicturePreviews();
		await previews.load(path.join(root, "shared", "photo.bmp"), root);
		assert.match(previews.get(path.join(root, "shared", "photo.bmp"))?.problem ?? "", /プロジェクトの外/);
		// A project opened through a link still references its own files in place.
		const inside = write(path.join(root, "images", "in.bmp"), data);
		assert.deepEqual(await planPictureAsset(path.join(base, "alias"), inside, data), { relativePath: "images/in.bmp" });
		assert.deepEqual(await planPictureAsset(root, path.join(base, "alias", "images", "in.bmp"), data), { relativePath: "images/in.bmp" });
		if (path.sep === "/") await assert.rejects(planPictureAsset(root, write(path.join(base, "linked", "a:b.bmp"), data), data), PictureError);
	});

	it("never overwrites an existing asset", async () => {
		const root = project("race");
		const data = bmp(2, 2);
		const plan = await planPictureAsset(root, write(path.join(base, "downloads", "race.bmp"), data), data);
		write(plan.copyTo!, "other");
		await assert.rejects(writePictureAsset(plan, data));
		assert.equal(fs.readFileSync(plan.copyTo!, "utf8"), "other");
	});

	it("reads picture files with the same checks", async () => {
		const good = write(path.join(base, "read", "good.jpg"), jpeg(app0, frame(0xc0, 24, 16)));
		assert.deepEqual((await readPictureFile(good)).info, { format: "jpeg", width: 24, height: 16, converted: false });
		await assert.rejects(readPictureFile(path.join(base, "read", "missing.bmp")), PictureError);
		await assert.rejects(readPictureFile(path.join(base, "read")), PictureError);
		await assert.rejects(readPictureFile(write(path.join(base, "read", "empty.bmp"))), /16 MiB/);
		await assert.rejects(readPictureFile(write(path.join(base, "read", "text.bmp"), "hello")), /判別できません/);
		// The format comes from the content, not from the extension.
		assert.deepEqual((await readPictureFile(write(path.join(base, "read", "really-png.bmp"), png(8, 4)))).info, { format: "png", width: 8, height: 4, converted: true });
	});

	it("previews converted formats with their MIME type, and TIFF not at all", async () => {
		const root = project("formats");
		const picture = png(8, 4);
		const tiff = Buffer.from([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0]);
		write(path.join(root, "forms", "assets", "icon.png"), picture);
		write(path.join(root, "forms", "assets", "scan.tif"), tiff);
		const parsed = parseSpec(
			formatJson({
				...newFormSpec("Main"),
				controls: [
					{ id: "a", name: "A", type: "Image", picture: { path: "forms/assets/icon.png" } },
					{ id: "b", name: "B", type: "Image", picture: { path: "forms/assets/scan.tif" } },
				],
			}),
		);
		assert.ok(parsed.ok);
		const previews = new PicturePreviews();
		await previews.refresh(pictureFiles(parsed.spec, root), root);
		// A format the webview cannot draw is neither previewed nor reported.
		assert.deepEqual(picturePreviewState(parsed.spec, root, previews), { dataUris: [pictureDataUri(picture, "png"), undefined], issues: [] });
		assert.ok(pictureDataUri(picture, "png").startsWith("data:image/png;base64,iVBOR"));
	});

	it("previews the pictures of a form and reports the ones it cannot show", async () => {
		const root = project("preview");
		const logo = write(path.join(root, "forms", "assets", "logo.bmp"), bmp(2, 2));
		write(path.join(root, "forms", "assets", "broken.bmp"), "not a picture");
		const image = (id: string, picture: unknown) => ({ id, name: id, type: "Image", picture });
		const parsed = parseSpec(
			formatJson({
				...newFormSpec("Main"),
				controls: [
					{ id: "l", name: "L", type: "Label" },
					image("a", { path: "forms/assets/logo.bmp" }),
					image("b", { path: "forms/assets/missing.bmp" }),
					image("c", { path: "forms/assets/broken.bmp" }),
					image("d", { path: "../outside.bmp" }),
					image("e", { remove: true }),
					image("f", { path: "forms/assets/logo.bmp" }),
				],
			}),
		);
		assert.ok(parsed.ok);
		const files = pictureFiles(parsed.spec, root);
		assert.deepEqual(files, ["logo.bmp", "missing.bmp", "broken.bmp"].map((name) => path.join(root, "forms", "assets", name)));
		const previews = new PicturePreviews();
		// Nothing is loaded yet: no previews and no file problems.
		assert.deepEqual(picturePreviewState(parsed.spec, root, previews).issues.map((i) => i.field), ["controls[4].picture.path"]);
		assert.equal(await previews.refresh(files, root), true);
		assert.equal(await previews.refresh(files, root), false);

		const state = picturePreviewState(parsed.spec, root, previews);
		const uri = pictureDataUri(bmp(2, 2), "bmp");
		assert.deepEqual(state.dataUris, [undefined, uri, undefined, undefined, undefined, undefined, uri]);
		assert.deepEqual(state.issues.map((i) => [i.code, i.severity, i.field]), [
			["UFP001", "warning", "controls[2].picture.path"],
			["UFP001", "warning", "controls[3].picture.path"],
			["UFP001", "warning", "controls[4].picture.path"],
		]);
		assert.match(state.issues[0].message, /見つからない/);
		assert.match(state.issues[1].message, /判別できません/);
		const [problem] = problems(parsed.spec, state.issues);
		assert.match(parsed.spec.text.split("\n")[problem.range!.start.line], /missing\.bmp/);

		// A file that changed on disk is read again; files no longer used are dropped.
		fs.writeFileSync(logo, bmp(4, 4));
		assert.equal(await previews.refresh(files, root), true);
		assert.equal(previews.get(logo)?.dataUri, pictureDataUri(bmp(4, 4), "bmp"));
		write(path.join(root, "forms", "assets", "missing.bmp"), bmp(1, 1));
		assert.equal(await previews.refresh(files.slice(1), root), true);
		assert.equal(previews.get(logo), undefined);
		assert.equal(previews.get(files[1])?.problem, undefined);
	});

	it("checks mouse icons, and previews the pictures of other controls and of the form", async () => {
		// An icon (type 1) or cursor (type 2) with one 2×2 image.
		const icon = (type: number) => {
			const data = Buffer.alloc(22 + 40, 0);
			data.writeUInt16LE(type, 2);
			data.writeUInt16LE(1, 4);
			data.writeUInt8(2, 6);
			data.writeUInt8(2, 7);
			data.writeUInt32LE(40, 14);
			data.writeUInt32LE(22, 18);
			return data;
		};
		assert.deepEqual(inspectIcon(icon(1)), { format: "ico", width: 2, height: 2, converted: false });
		assert.deepEqual(inspectIcon(icon(2)), { format: "cur", width: 2, height: 2, converted: false });
		// A bitmap MS Forms holds as it is will do; one that would be converted will not.
		assert.equal(inspectIcon(bmp(2, 2)).format, "bmp");
		assert.throws(() => inspectIcon(Buffer.from("\x89PNG\r\n\x1a\n", "latin1")), PictureError);
		assert.throws(() => inspectIcon(icon(1).subarray(0, 30)), PictureError);
		assert.deepEqual(ICON_EXTENSIONS, ["ico", "cur"]);

		const root = project("others");
		write(path.join(root, "forms", "assets", "logo.bmp"), bmp(2, 2));
		write(path.join(root, "forms", "assets", "hand.cur"), icon(2));
		const parsed = parseSpec(
			formatJson({
				...newFormSpec("Main"),
				form: { name: "Main", picture: { path: "forms/assets/logo.bmp" }, mouseIcon: { path: "forms/assets/hand.cur" } },
				controls: [
					{ id: "l", name: "L", type: "Label", picture: { path: "forms/assets/logo.bmp" }, mouseIcon: { path: "forms/assets/logo.bmp" } },
					{ id: "b", name: "B", type: "CommandButton", mouseIcon: { path: "forms/assets/logo.png" } },
					{ id: "t", name: "T", type: "TextBox", picture: { path: "forms/assets/ignored.bmp" } },
				],
			}),
		);
		assert.ok(parsed.ok);
		const spec = parsed.spec;
		const files = pictureFiles(spec, root);
		// The TextBox has no picture: its path is not read.
		assert.deepEqual(files.sort(), ["hand.cur", "logo.bmp", "logo.png"].map((name) => path.join(root, "forms", "assets", name)).sort());
		const previews = new PicturePreviews();
		await previews.refresh(files, root);
		const state = picturePreviewState(spec, root, previews);
		const uri = pictureDataUri(bmp(2, 2), "bmp");
		assert.equal(state.form, uri);
		assert.deepEqual(state.dataUris, [uri, undefined, undefined]);
		// A mouse icon is checked as an icon, and not previewed.
		assert.deepEqual(state.issues.map((i) => i.field), ["controls[1].mouseIcon.path"]);
	});
});

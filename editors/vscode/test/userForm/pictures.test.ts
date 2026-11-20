import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { after, describe, it } from "node:test";
import { formatJson } from "../../src/userForm/jsonSource";
import {
	findProjectRoot,
	inspectPicture,
	MAX_PICTURE_BYTES,
	pictureDataUri,
	PictureError,
	pictureFiles,
	PicturePreviews,
	picturePreviewState,
	planPictureAsset,
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

const rejects = (data: Uint8Array, pattern: RegExp) =>
	assert.throws(
		() => inspectPicture(data),
		(error: unknown) => error instanceof PictureError && pattern.test(error.message),
	);

describe("picture file validation", () => {
	it("accepts uncompressed 24-bit and 32-bit BMP files", () => {
		assert.deepEqual(inspectPicture(bmp(3, 2)), { format: "bmp", width: 3, height: 2 });
		assert.deepEqual(inspectPicture(bmp(5, -4, 32)), { format: "bmp", width: 5, height: 4 });
		// Bytes after the pixel array and a later pixel offset are fine.
		const padded = Buffer.concat([bmp(2, 2), Buffer.alloc(8)]);
		assert.equal(inspectPicture(padded).format, "bmp");
		assert.ok(pictureDataUri(bmp(1, 1), "bmp").startsWith("data:image/bmp;base64,Qk"));
	});

	it("rejects BMP files the compiler cannot embed", () => {
		rejects(bmp(3, 2).subarray(0, 40), /ヘッダー/);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt32LE(124, 14)), /BITMAPINFOHEADER/);
		rejects(bmp(3, 2, 24, (d) => d.writeInt32LE(0, 18)), /幅と高さ/);
		rejects(bmp(3, 2, 24, (d) => d.writeInt32LE(-3, 18)), /幅と高さ/);
		rejects(bmp(3, 2, 24, (d) => d.writeInt32LE(0, 22)), /幅と高さ/);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt16LE(0, 26)), /プレーン/);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt16LE(8, 28)), /8 ビット/);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt32LE(1, 30)), /無圧縮/);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt32LE(40, 10)), /ピクセル データ/);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt32LE(58, 10)), /ピクセル データ/);
		rejects(bmp(3, 2).subarray(0, 60), /ピクセル データ/);
		rejects(bmp(3, 2, 24, (d) => d.writeInt32LE(-0x80000000, 22)), /幅と高さ/);
		// Sizes in the header may be 0, but must not contradict the pixel data.
		assert.equal(inspectPicture(bmp(3, 2, 24, (d) => d.writeUInt32LE(0, 2))).width, 3);
		assert.equal(inspectPicture(bmp(3, 2, 24, (d) => d.writeUInt32LE(24, 34))).width, 3);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt32LE(d.length + 1, 2)), /ファイル サイズ/);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt32LE(60, 2)), /ファイル サイズ/);
		rejects(bmp(3, 2, 24, (d) => d.writeUInt32LE(23, 34)), /画像サイズ/);
		// A header that claims more pixels than the limit, without the data for them.
		rejects(bmp(1, 1, 24, (d) => (d.writeInt32LE(5000, 18), d.writeInt32LE(4000, 22))), /16,000,000/);
	});

	it("reads JPEG dimensions from the first frame header", () => {
		const table = segment(0xc4, [0, 1, 2, 3]);
		assert.deepEqual(inspectPicture(jpeg(app0, frame(0xc0, 24, 16))), { format: "jpeg", width: 24, height: 16 });
		// DHT (0xC4) is not a frame header; progressive frames (0xC2) count.
		assert.deepEqual(inspectPicture(jpeg(app0, table, frame(0xc2, 640, 480), table)), { format: "jpeg", width: 640, height: 480 });
		// Fill bytes before a marker are skipped.
		assert.deepEqual(inspectPicture(jpeg(app0, [0xff, 0xff], frame(0xc1, 300, 2))), { format: "jpeg", width: 300, height: 2 });
		assert.deepEqual(inspectPicture(jpeg(frame(0xc0, 4000, 4000))), { format: "jpeg", width: 4000, height: 4000 });
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

	it("rejects other formats and sizes outside 1 byte to 16 MiB", () => {
		rejects(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), /BMP または JPEG/);
		rejects(Buffer.from("B"), /BMP または JPEG/);
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
		assert.deepEqual((await readPictureFile(good)).info, { format: "jpeg", width: 24, height: 16 });
		await assert.rejects(readPictureFile(path.join(base, "read", "missing.bmp")), PictureError);
		await assert.rejects(readPictureFile(path.join(base, "read")), PictureError);
		await assert.rejects(readPictureFile(write(path.join(base, "read", "empty.bmp"))), /16 MiB/);
		await assert.rejects(readPictureFile(write(path.join(base, "read", "text.bmp"), "hello")), /BMP または JPEG/);
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
		assert.match(state.issues[1].message, /BMP または JPEG/);
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
});

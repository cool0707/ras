// The files of the page are loaded with the id of the build (`app.js?v=<id>`): a browser that
// kept an earlier page's file in its cache would otherwise mix it with the new ones (the
// Designer of one build talking to the page of another).

/** `file` with the build's id. */
export const versioned = (file: string): string => (typeof RAS_BUILD === "string" && RAS_BUILD ? `${file}?v=${RAS_BUILD}` : file);

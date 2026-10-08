// Clean, minimal stream shim for jszip in browser/obsidian
"use strict";

class Stream {}
class Readable extends Stream {}
class Writable extends Stream {}
class Duplex extends Stream {}
class Transform extends Stream {}
class PassThrough extends Stream {}

module.exports = {
  Stream,
  Readable,
  Writable,
  Duplex,
  Transform,
  PassThrough,
};

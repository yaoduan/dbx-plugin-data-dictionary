/**
 * Minimal ZIP writer for the data dictionary exporter.
 *
 * OOXML files (.docx / .xlsx) are ZIP containers, and the plugin sandbox has no
 * network access and no bundler, so the archive is built here instead of
 * pulling in a library. Entries are deflated with the platform
 * `CompressionStream("deflate-raw")` when it exists and stored uncompressed
 * otherwise -- both are valid ZIP, so Word and Excel read either one.
 */
(function (global) {
  "use strict";

  var CRC_TABLE = (function () {
    var table = new Int32Array(256);
    for (var i = 0; i < 256; i += 1) {
      var value = i;
      for (var bit = 0; bit < 8; bit += 1) {
        value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
      }
      table[i] = value;
    }
    return table;
  })();

  function crc32(bytes) {
    var crc = -1;
    for (var i = 0; i < bytes.length; i += 1) {
      crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ bytes[i]) & 0xff];
    }
    return (crc ^ -1) >>> 0;
  }

  var textEncoder = new TextEncoder();

  function Writer(capacity) {
    this.bytes = new Uint8Array(capacity > 0 ? capacity : 65536);
    this.length = 0;
  }

  Writer.prototype.ensure = function (extra) {
    if (this.length + extra <= this.bytes.length) return;
    var size = this.bytes.length;
    while (size < this.length + extra) size *= 2;
    var grown = new Uint8Array(size);
    grown.set(this.bytes.subarray(0, this.length), 0);
    this.bytes = grown;
  };

  Writer.prototype.u16 = function (value) {
    this.ensure(2);
    this.bytes[this.length] = value & 0xff;
    this.bytes[this.length + 1] = (value >>> 8) & 0xff;
    this.length += 2;
  };

  Writer.prototype.u32 = function (value) {
    this.ensure(4);
    this.bytes[this.length] = value & 0xff;
    this.bytes[this.length + 1] = (value >>> 8) & 0xff;
    this.bytes[this.length + 2] = (value >>> 16) & 0xff;
    this.bytes[this.length + 3] = (value >>> 24) & 0xff;
    this.length += 4;
  };

  Writer.prototype.raw = function (bytes) {
    this.ensure(bytes.length);
    this.bytes.set(bytes, this.length);
    this.length += bytes.length;
  };

  Writer.prototype.result = function () {
    return this.bytes.subarray(0, this.length);
  };

  /** DOS date/time pair used by the local and central headers. */
  function dosDateTime(date) {
    var year = date.getFullYear();
    if (year < 1980) return { date: 33, time: 0 };
    return {
      date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
      time: (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1),
    };
  }

  /**
   * Raw-deflate `bytes`, or `null` when the host webview has no
   * CompressionStream. A null answer is not an error: the caller stores the
   * entry instead, which keeps the archive valid on older webviews.
   */
  async function deflateRaw(bytes) {
    if (typeof global.CompressionStream !== "function") return null;
    try {
      var stream = new global.CompressionStream("deflate-raw");
      var writer = stream.writable.getWriter();
      // Errors on either side surface through the read loop below.
      writer.write(bytes).catch(function () {});
      writer.close().catch(function () {});
      var reader = stream.readable.getReader();
      var chunks = [];
      var total = 0;
      for (;;) {
        var step = await reader.read();
        if (step.done) break;
        var chunk = step.value instanceof Uint8Array ? step.value : new Uint8Array(step.value);
        chunks.push(chunk);
        total += chunk.length;
      }
      var deflated = new Uint8Array(total);
      var offset = 0;
      for (var i = 0; i < chunks.length; i += 1) {
        deflated.set(chunks[i], offset);
        offset += chunks[i].length;
      }
      return deflated;
    } catch (error) {
      return null;
    }
  }

  function toBytes(content) {
    if (content instanceof Uint8Array) return content;
    if (content instanceof ArrayBuffer) return new Uint8Array(content);
    return textEncoder.encode(String(content));
  }

  /**
   * Build a ZIP archive from `[{ name, content }]`, where content is a string,
   * Uint8Array or ArrayBuffer. Returns a Uint8Array.
   */
  async function createZip(entries) {
    var stamp = dosDateTime(new Date());
    var body = new Writer(262144);
    var directory = new Writer(16384);
    var count = 0;

    for (var i = 0; i < entries.length; i += 1) {
      var entry = entries[i];
      var nameBytes = textEncoder.encode(entry.name);
      var data = toBytes(entry.content);
      var checksum = crc32(data);
      var compressed = data.length > 0 ? await deflateRaw(data) : null;
      // Deflate can grow incompressible payloads; store those instead.
      var useDeflate = compressed !== null && compressed.length < data.length;
      var payload = useDeflate ? compressed : data;
      var method = useDeflate ? 8 : 0;
      var offset = body.length;

      body.u32(0x04034b50);
      body.u16(useDeflate ? 20 : 10);
      body.u16(0x0800); // UTF-8 names
      body.u16(method);
      body.u16(stamp.time);
      body.u16(stamp.date);
      body.u32(checksum);
      body.u32(payload.length);
      body.u32(data.length);
      body.u16(nameBytes.length);
      body.u16(0);
      body.raw(nameBytes);
      body.raw(payload);

      directory.u32(0x02014b50);
      directory.u16(0x031e); // made by UNIX, zip 3.0
      directory.u16(useDeflate ? 20 : 10);
      directory.u16(0x0800);
      directory.u16(method);
      directory.u16(stamp.time);
      directory.u16(stamp.date);
      directory.u32(checksum);
      directory.u32(payload.length);
      directory.u32(data.length);
      directory.u16(nameBytes.length);
      directory.u16(0);
      directory.u16(0);
      directory.u16(0);
      directory.u16(0);
      directory.u32(0);
      directory.u32(offset);
      directory.raw(nameBytes);
      count += 1;
    }

    var directoryBytes = directory.result();
    var directoryOffset = body.length;
    body.raw(directoryBytes);
    body.u32(0x06054b50);
    body.u16(0);
    body.u16(0);
    body.u16(count);
    body.u16(count);
    body.u32(directoryBytes.length);
    body.u32(directoryOffset);
    body.u16(0);
    return body.result();
  }

  global.DictZip = { createZip: createZip, crc32: crc32 };
})(window);

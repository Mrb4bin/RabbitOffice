// Minimal dependency-free WebSocket server (RFC 6455) used by the Linux server and the desktop host.
'use strict';
const crypto = require('crypto');
const { EventEmitter } = require('events');

const MAX_PAYLOAD = 32 * 1024 * 1024;

class Socket extends EventEmitter {
  constructor(sock) {
    super(); this.sock = sock; this.buf = Buffer.alloc(0); this.frags = []; this.open = true;
    sock.setNoDelay(true);
    sock.on('data', (d) => this._data(d));
    sock.on('close', () => this._closed());
    sock.on('error', () => this._closed());
    this.ping = setInterval(() => this._frame(0x9, Buffer.alloc(0)), 25000);
  }
  _closed() { if (!this.open) return; this.open = false; clearInterval(this.ping); this.emit('close'); }
  _data(d) {
    this.buf = Buffer.concat([this.buf, d]);
    for (;;) {
      if (this.buf.length < 2) return;
      const b0 = this.buf[0], b1 = this.buf[1]; const fin = b0 & 0x80, op = b0 & 0x0f, masked = b1 & 0x80;
      let len = b1 & 0x7f, off = 2;
      if (len === 126) { if (this.buf.length < 4) return; len = this.buf.readUInt16BE(2); off = 4; }
      else if (len === 127) { if (this.buf.length < 10) return; len = Number(this.buf.readBigUInt64BE(2)); off = 10; }
      if (len > MAX_PAYLOAD) return this.close(1009);
      const mlen = masked ? 4 : 0; if (this.buf.length < off + mlen + len) return;
      const mask = masked ? this.buf.subarray(off, off + 4) : null;
      const payload = Buffer.from(this.buf.subarray(off + mlen, off + mlen + len));
      if (mask) for (let i = 0; i < payload.length; i++) payload[i] ^= mask[i & 3];
      this.buf = this.buf.subarray(off + mlen + len);
      if (op === 0x8) { this.close(); return; }
      if (op === 0x9) { this._frame(0xA, payload); continue; }
      if (op === 0xA) continue;
      if (op === 0x1 || op === 0x2 || op === 0x0) { this.frags.push(payload); if (fin) { const msg = Buffer.concat(this.frags); this.frags = []; this.emit('message', msg.toString('utf8')); } }
    }
  }
  _frame(op, payload) {
    if (!this.open) return;
    const len = payload.length; let head;
    if (len < 126) head = Buffer.from([0x80 | op, len]);
    else if (len < 65536) { head = Buffer.alloc(4); head[0] = 0x80 | op; head[1] = 126; head.writeUInt16BE(len, 2); }
    else { head = Buffer.alloc(10); head[0] = 0x80 | op; head[1] = 127; head.writeBigUInt64BE(BigInt(len), 2); }
    try { this.sock.write(Buffer.concat([head, payload])); } catch { this._closed(); }
  }
  send(str) { this._frame(0x1, Buffer.from(str, 'utf8')); }
  close(code = 1000) { const b = Buffer.alloc(2); b.writeUInt16BE(code); this._frame(0x8, b); try { this.sock.end(); } catch { /* ignore */ } this._closed(); }
}

// Accept an HTTP upgrade request and return a Socket (or null if it is not a websocket upgrade).
function accept(req, sock) {
  const key = req.headers['sec-websocket-key'];
  if (!key || String(req.headers.upgrade).toLowerCase() !== 'websocket') { sock.destroy(); return null; }
  const acceptKey = crypto.createHash('sha1').update(key + '258EAFA5-E914-47DA-95CA-C5AB0DC85B11').digest('base64');
  sock.write(`HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Accept: ${acceptKey}\r\n\r\n`);
  return new Socket(sock);
}
module.exports = { accept, Socket };

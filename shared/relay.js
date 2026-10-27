// Board collaboration room: last-writer-wins object store + presence relay.
// Shared by Rabbit Office Server (boards on the server) and the desktop app (direct / LAN host).
'use strict';
const crypto = require('crypto');

class Room {
  constructor(id, { objects = [], meta = {}, pin = '', onChange } = {}) {
    this.id = id; this.objects = new Map(objects.map((o) => [o.id, o])); this.meta = meta; this.pin = pin;
    this.clients = new Map(); this.onChange = onChange; this.dirty = false;
  }
  join(ws, user = null) {
    let peer = null;
    ws.on('message', (raw) => {
      let m; try { m = JSON.parse(raw); } catch { return; }
      if (!peer) {
        if (m.t !== 'hello') return;
        if (this.pin && String(m.pin || '') !== this.pin) { ws.send(JSON.stringify({ t: 'error', message: 'Неверный PIN-код сессии' })); return ws.close(4003); }
        peer = { id: crypto.randomBytes(6).toString('hex'), name: String(user?.displayName || m.name || 'Гость').slice(0, 40), color: /^#[0-9a-f]{6}$/i.test(m.color) ? m.color : '#7b4fd6' };
        this.clients.set(ws, peer);
        ws.send(JSON.stringify({ t: 'welcome', you: peer, objects: [...this.objects.values()].filter((o) => !o.deleted), meta: this.meta, peers: [...this.clients.values()].filter((p) => p !== peer) }));
        this.broadcast({ t: 'peer-join', peer }, ws);
        return;
      }
      if (m.t === 'ops' && Array.isArray(m.ops)) {
        const accepted = [];
        for (const o of m.ops.slice(0, 5000)) {
          if (!o || typeof o.id !== 'string' || o.id.length > 64) continue;
          const cur = this.objects.get(o.id);
          if (cur && (cur.v || 0) > (o.v || 0)) continue;
          this.objects.set(o.id, o); accepted.push(o);
        }
        if (accepted.length) { this.broadcast({ t: 'ops', ops: accepted }, ws); this.dirty = true; this.onChange?.(this); }
      } else if (m.t === 'cursor') this.broadcast({ t: 'cursor', id: peer.id, x: m.x, y: m.y }, ws);
      else if (m.t === 'meta' && m.meta && typeof m.meta === 'object') { Object.assign(this.meta, m.meta); this.broadcast({ t: 'meta', meta: this.meta }, ws); this.onChange?.(this); }
    });
    ws.on('close', () => { const p = this.clients.get(ws); this.clients.delete(ws); if (p) this.broadcast({ t: 'peer-leave', id: p.id }); });
  }
  broadcast(msg, except) { const s = JSON.stringify(msg); for (const ws of this.clients.keys()) if (ws !== except) ws.send(s); }
  snapshot() { return { objects: [...this.objects.values()].filter((o) => !o.deleted), meta: this.meta }; }
  get online() { return this.clients.size; }
}
module.exports = { Room };

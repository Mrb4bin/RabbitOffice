#!/usr/bin/env node
// Rabbit Office Server CLI: setup wizard, user management, config.
'use strict';
const readline = require('readline');
const crypto = require('crypto');
const store = require('../src/store.js');

const [cmd, sub, ...args] = process.argv.slice(2);
const rl = () => readline.createInterface({ input: process.stdin, output: process.stdout });
const ask = (q, def = '') => new Promise((res) => { const r = rl(); r.question(`${q}${def ? ` [${def}]` : ''}: `, (a) => { r.close(); res(a.trim() || def); }); });
const askHidden = (q) => new Promise((res) => { const r = rl(); r.stdoutMuted = true; r._writeToOutput = (s) => { if (!r.stdoutMuted || s.includes(q)) r.output.write(s); }; r.question(q + ': ', (a) => { r.close(); process.stdout.write('\n'); res(a); }); });
const gen = () => crypto.randomBytes(9).toString('base64url');

async function setup() {
  store.init();
  const cfg = store.loadConfig();
  console.log('\n  🐇 Мастер настройки Rabbit Office Server\n');
  cfg.serverName = await ask('Название сервера', cfg.serverName);
  console.log('\n  Режим работы:\n   1) Хранилище файлов\n   2) Доски для совместной работы\n   3) Хранилище и доски');
  const m = await ask('Выберите 1, 2 или 3', { storage: '1', board: '2', both: '3' }[cfg.mode]);
  cfg.mode = { 1: 'storage', 2: 'board', 3: 'both' }[m] || 'both';
  console.log('\n  Домен: если указать домен, сервер сам получит бесплатный сертификат Let\'s Encrypt\n  и будет работать по HTTPS на портах 80/443. Оставьте пустым для работы по IP.');
  cfg.domain = (await ask('Домен (например office.example.com)', cfg.domain)).toLowerCase();
  if (cfg.domain) cfg.email = await ask('E-mail для уведомлений Let\'s Encrypt', cfg.email);
  else cfg.port = +(await ask('Порт', String(cfg.port))) || 8080;
  store.saveConfig(cfg);
  if (!store.listUsers().some((u) => u.role === 'admin')) {
    const user = (await ask('Логин администратора', 'admin')).toLowerCase();
    let pass = await askHidden('Пароль администратора (Enter — сгенерировать)');
    const generated = !pass; if (!pass) pass = gen();
    store.createUser({ username: user, password: pass, role: 'admin', displayName: 'Администратор' });
    console.log(`\n  ✔ Администратор создан: ${user}${generated ? ` / пароль: ${pass}` : ''}`);
  }
  const url = cfg.domain ? `https://${cfg.domain}` : `http://<IP-сервера>:${cfg.port}`;
  console.log(`\n  ✔ Настройка сохранена.\n\n  Офис в браузере:   ${url}\n  Админ-панель:      ${url}/admin\n  В приложениях:     «Подключить сервер» → ${cfg.domain || '<IP>:' + cfg.port}\n`);
}

async function main() {
  switch (cmd) {
    case 'setup': return setup();
    case 'start': return require('../src/server.js').start();
    case 'user': {
      store.init();
      if (sub === 'list') { for (const u of store.listUsers()) console.log(`${u.username.padEnd(20)} ${u.role.padEnd(6)} ${u.disabled ? 'отключён' : ''} ${u.displayName}`); return; }
      if (sub === 'add') { const name = args[0] || await ask('Логин'); const pass = args[1] || gen(); store.createUser({ username: name, password: pass, role: args.includes('--admin') ? 'admin' : 'user', displayName: await ask('Имя', name) }); console.log(`✔ Пользователь ${name} создан. Пароль: ${pass}`); return; }
      if (sub === 'passwd') { const name = args[0] || await ask('Логин'); const pass = args[1] || gen(); store.setPassword(name, pass); console.log(`✔ Новый пароль для ${name}: ${pass}`); return; }
      if (sub === 'del') { store.deleteUser(args[0]); console.log('✔ Удалён'); return; }
      break;
    }
    case 'backup': {
      const { execFileSync } = require('child_process'); const path = require('path');
      const out = path.resolve(sub || `rabbit-office-backup-${new Date().toISOString().slice(0, 10)}.tar.gz`);
      execFileSync('tar', ['-czf', out, '--exclude=./certs', '-C', store.dataDir(), '.'], { stdio: 'inherit' });
      console.log(`✔ Резервная копия: ${out}`); return;
    }
    case 'restore': {
      if (!sub) { console.log('Укажите файл: rabbit-office restore <архив.tar.gz>'); return; }
      const ok = (await ask(`Текущие данные в ${store.dataDir()} будут заменены. Продолжить? (да/нет)`, 'нет')).toLowerCase();
      if (!['да', 'y', 'yes', 'д'].includes(ok)) return;
      const { execFileSync } = require('child_process'); const fs = require('fs'); const path = require('path');
      const keep = path.join(store.dataDir(), 'certs');
      for (const e of fs.readdirSync(store.dataDir())) if (path.join(store.dataDir(), e) !== keep) fs.rmSync(path.join(store.dataDir(), e), { recursive: true, force: true });
      execFileSync('tar', ['-xzf', path.resolve(sub), '-C', store.dataDir()], { stdio: 'inherit' });
      console.log('✔ Данные восстановлены. Перезапустите службу: sudo systemctl restart rabbit-office'); return;
    }
    case 'config': {
      const cfg = store.loadConfig();
      if (sub === 'get' || !sub) { console.log(JSON.stringify(cfg, null, 2)); return; }
      if (sub === 'set') { const [k, v] = args; cfg[k] = v === 'true' ? true : v === 'false' ? false : isFinite(+v) && v !== '' ? +v : v; store.saveConfig(cfg); console.log(`✔ ${k} = ${cfg[k]} (перезапустите службу)`); return; }
      break;
    }
  }
  console.log(`Rabbit Office Server
  rabbit-office setup                  мастер настройки
  rabbit-office start                  запустить сервер (обычно через systemd)
  rabbit-office user list|add|passwd|del [логин] [пароль] [--admin]
  rabbit-office config get | set <ключ> <значение>
  rabbit-office backup [файл.tar.gz]   резервная копия данных
  rabbit-office restore <файл.tar.gz>  восстановление из копии`);
}
main().catch((e) => { console.error('Ошибка:', e.message); process.exit(1); });

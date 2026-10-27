// Template catalog (metadata + thumbnails). Content lives in each app.
const line = (w, top, h = 3, c = '#d6d9e0', l = 12) => `<i style="position:absolute;left:${l}%;top:${top}%;width:${w}%;height:${h}px;background:${c};border-radius:2px"></i>`;
const docThumb = (title, accent) => `<div style="position:absolute;inset:8% 22%;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.12)">${title ? line(50, 12, 5, accent) : ''}${[24, 32, 40, 48, 56, 64, 72].map((t, i) => line(i % 3 === 2 ? 50 : 76, t)).join('')}</div>`;
const sheetThumb = (accent) => `<div style="position:absolute;inset:0;background:#fff;background-image:linear-gradient(#e3e6ea 1px,transparent 1px),linear-gradient(90deg,#e3e6ea 1px,transparent 1px);background-size:100% 14%,20% 100%">${line(100, 0, 14, accent + '33', 0)}</div>`;
const slideThumb = (bg, fg, accent) => `<div style="position:absolute;inset:0;background:${bg}">${line(56, 34, 9, fg, 10)}${line(36, 52, 4, accent, 10)}</div>`;
const deskThumb = (kind) => {
  const dots = 'background-color:#fafafa;background-image:radial-gradient(#d4d4d8 1px,transparent 1px);background-size:10px 10px';
  if (kind === 'kanban') return `<div style="position:absolute;inset:0;${dots}">${[8, 38, 68].map((l, i) => `<i style="position:absolute;left:${l}%;top:12%;width:26%;height:76%;background:#eef0f4;border-radius:4px"></i>${[22, 42].map((t) => `<i style="position:absolute;left:${l + 3}%;top:${t}%;width:20%;height:16%;background:${['#ffe58a', '#a8e6b8', '#a8d4ff'][i]};border-radius:2px"></i>`).join('')}`).join('')}</div>`;
  if (kind === 'retro') return `<div style="position:absolute;inset:0;${dots}">${['#a8e6b8', '#ffb3b3', '#a8d4ff'].map((c, i) => `<i style="position:absolute;left:${8 + i * 30}%;top:14%;width:26%;height:72%;border:2px solid ${c};border-radius:6px"></i>`).join('')}</div>`;
  if (kind === 'brainstorm') return `<div style="position:absolute;inset:0;${dots}"><i style="position:absolute;left:40%;top:38%;width:20%;height:24%;background:#c4b5fd;border-radius:50%"></i>${[[12, 14, '#ffe58a'], [70, 12, '#a8e6b8'], [14, 66, '#ffb3b3'], [70, 64, '#a8d4ff']].map(([l, t, c]) => `<i style="position:absolute;left:${l}%;top:${t}%;width:18%;height:22%;background:${c};border-radius:2px"></i>`).join('')}</div>`;
  return `<div style="position:absolute;inset:0;${dots}"></div>`;
};

export const TEMPLATES = {
  docs: [
    { id: '', name: 'Новый документ', desc: 'Пустая страница', thumb: docThumb(false) },
    { id: 'letter', name: 'Деловое письмо', desc: 'Шапка, реквизиты, подпись', thumb: docThumb(true, '#2f62c9') },
    { id: 'report', name: 'Отчёт', desc: 'Титул, разделы, таблица', thumb: docThumb(true, '#1d8a52') },
    { id: 'resume', name: 'Резюме', desc: 'Опыт, навыки, контакты', thumb: docThumb(true, '#7b4fd6') },
    { id: 'notes', name: 'Протокол встречи', desc: 'Повестка, решения, задачи', thumb: docThumb(true, '#d2552d') },
  ],
  xel: [
    { id: '', name: 'Пустая книга', desc: 'Чистый лист', thumb: sheetThumb('#1d8a52') },
    { id: 'budget', name: 'Личный бюджет', desc: 'Доходы, расходы, итоги', thumb: sheetThumb('#2f62c9') },
    { id: 'invoice', name: 'Счёт', desc: 'Позиции, НДС, сумма', thumb: sheetThumb('#d2552d') },
    { id: 'schedule', name: 'План проекта', desc: 'Задачи, сроки, статус', thumb: sheetThumb('#7b4fd6') },
  ],
  point: [
    { id: '', name: 'Пустая презентация', desc: '16:9', thumb: slideThumb('#ffffff', '#30343b', '#d2552d') },
    { id: 'midnight', name: 'Полночь', desc: 'Тёмная тема', thumb: slideThumb('#151a2d', '#ffffff', '#6aa7ff') },
    { id: 'pitch', name: 'Питч стартапа', desc: '8 слайдов', thumb: slideThumb('linear-gradient(135deg,#ff8a5b,#d2552d)', '#ffffff', '#ffe0d1') },
    { id: 'minimal', name: 'Минимализм', desc: 'Светлая, крупная типографика', thumb: slideThumb('#f5f1ea', '#1f1f1f', '#1d8a52') },
  ],
  desk: [
    { id: '', name: 'Пустая доска', desc: 'Бесконечный холст', thumb: deskThumb('') },
    { id: 'brainstorm', name: 'Мозговой штурм', desc: 'Идеи вокруг темы', thumb: deskThumb('brainstorm') },
    { id: 'kanban', name: 'Канбан', desc: 'Сделать / В работе / Готово', thumb: deskThumb('kanban') },
    { id: 'retro', name: 'Ретроспектива', desc: 'Хорошо / Плохо / Улучшить', thumb: deskThumb('retro') },
  ],
};

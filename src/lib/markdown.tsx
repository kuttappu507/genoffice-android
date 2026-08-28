import type { ReactElement } from 'react';

function inline(text: string, keyBase: string): ReactElement[] {
  const parts = text.split(/(\*\*[^*]+\*\*|`[^`]+`)/g);
  return parts
    .filter((p) => p !== '')
    .map((p, n) => {
      const key = `${keyBase}-${n}`;
      if (p.startsWith('**') && p.endsWith('**')) return <strong key={key}>{p.slice(2, -2)}</strong>;
      if (p.startsWith('`') && p.endsWith('`')) return <code key={key}>{p.slice(1, -1)}</code>;
      return <span key={key}>{p}</span>;
    });
}

const BLOCK_START = /^(#{1,4}\s|[-*]\s|\d+\.\s|```)/;

/** Minimal markdown renderer: headings, lists, code fences, bold, inline code. */
export function Markdown({ text }: { text: string }): ReactElement {
  const blocks: ReactElement[] = [];
  const lines = text.split('\n');
  let i = 0;
  let key = 0;

  while (i < lines.length) {
    const line = lines[i];
    if (line.startsWith('```')) {
      const buf: string[] = [];
      i++;
      while (i < lines.length && !lines[i].startsWith('```')) {
        buf.push(lines[i]);
        i++;
      }
      i++;
      blocks.push(
        <pre key={key++}>
          <code>{buf.join('\n')}</code>
        </pre>,
      );
      continue;
    }
    const h = /^(#{1,4})\s+(.*)/.exec(line);
    if (h) {
      const lvl = h[1].length;
      const content = inline(h[2], `h${key}`);
      if (lvl <= 2) blocks.push(<h3 key={key++}>{content}</h3>);
      else blocks.push(<h4 key={key++}>{content}</h4>);
      i++;
      continue;
    }
    if (/^[-*]\s+/.test(line)) {
      const items: string[] = [];
      while (i < lines.length && /^[-*]\s+/.test(lines[i])) {
        items.push(lines[i].replace(/^[-*]\s+/, ''));
        i++;
      }
      const k = key++;
      blocks.push(
        <ul key={k}>
          {items.map((it, n) => (
            <li key={n}>{inline(it, `li${k}-${n}`)}</li>
          ))}
        </ul>,
      );
      continue;
    }
    if (/^\d+\.\s+/.test(line)) {
      const items: string[] = [];
      while (i < lines.length && /^\d+\.\s+/.test(lines[i])) {
        items.push(lines[i].replace(/^\d+\.\s+/, ''));
        i++;
      }
      const k = key++;
      blocks.push(
        <ol key={k}>
          {items.map((it, n) => (
            <li key={n}>{inline(it, `ol${k}-${n}`)}</li>
          ))}
        </ol>,
      );
      continue;
    }
    if (!line.trim()) {
      i++;
      continue;
    }
    const para: string[] = [];
    while (i < lines.length && lines[i].trim() && !BLOCK_START.test(lines[i])) {
      para.push(lines[i]);
      i++;
    }
    blocks.push(<p key={key++}>{inline(para.join(' '), `p${key}`)}</p>);
  }

  return <div className="md">{blocks}</div>;
}

// Seeded mutations of XML text, shared by the fuzz and differential tests.

export function rng(seed: number): () => number {        // mulberry32
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const TOKENS = [
  '<', '>', '/>', '</', '&', ';', '&amp;', '&#x0;', '&#xD800;', '&#1114112;', '&#x1F600;', '&lt', '&nbsp;', '&constructor;',
  '"', "'", '=', ':', ' xmlns:p=""', ' xmlns=""', ' xmlns:q="urn:q"', ' q:x="1"', '<!--', '-->', '--', '<![CDATA[', ']]>',
  '<!DOCTYPE x [<!ENTITY a "b">]>', '<?pi x?>', '<?xml version="1.0"?>', '\0', '\u0007', '￾', '\uD800', '\uDC00',
  '😀', '\r', '\n', '\r\n', '\t', ' ', 'é', 'p:', ':a', 'a:b:c', '<a>', '</a>', '<p:sp>', '</p:sp>', ' a="1"', " b='2'",
];

export function mutate(src: string, rnd: () => number): string {
  let s = src;
  const edits = 1 + Math.floor(rnd() * 3);
  for (let k = 0; k < edits; k++) {
    const at = Math.floor(rnd() * (s.length + 1));
    const r = rnd();
    if (r < 0.3) s = s.slice(0, at) + s.slice(at + 1 + Math.floor(rnd() * 12));                       // delete a run
    else if (r < 0.65) s = s.slice(0, at) + TOKENS[Math.floor(rnd() * TOKENS.length)] + s.slice(at);    // insert a token
    else if (r < 0.85) {                                                                                 // copy a slice elsewhere
      const from = Math.floor(rnd() * s.length);
      s = s.slice(0, at) + s.slice(from, from + 1 + Math.floor(rnd() * 40)) + s.slice(at);
    } else if (r < 0.97) s = s.slice(0, at) + String.fromCharCode(Math.floor(rnd() * 128)) + s.slice(at + 1);   // overwrite a char
    else s = s.slice(0, at);                                                                             // truncate
  }
  return s;
}

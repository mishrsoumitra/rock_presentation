// apps/web/src/ui/bullets.ts  -- the bullet and number styles the ribbon offers, and the indent that goes with them
import type { BulletSpec, ParaPatch } from '@pptx/editor-core';
import { autonumberLabel } from '@pptx/text';

/** Bullet characters every text font has, so they look the same here and in PowerPoint. */
export const BULLET_STYLES: { char: string; name: string }[] = [
  { char: '•', name: 'Filled Round' }, { char: '○', name: 'Hollow Round' }, { char: '■', name: 'Filled Square' },
  { char: '□', name: 'Hollow Square' }, { char: '▪', name: 'Small Square' }, { char: '–', name: 'Dash' },
];

/** Numbering schemes (ST_TextAutonumberScheme), as PowerPoint's gallery orders them. */
export const NUMBER_STYLES: string[] = ['arabicPeriod', 'arabicParenR', 'romanUcPeriod', 'romanLcPeriod', 'alphaUcPeriod', 'alphaLcPeriod', 'alphaLcParenR'];

/** How a numbering scheme looks, for its menu item: "1. 2. 3.", "a) b) c)". */
export const numberSample = (scheme: string) => [1, 2, 3].map(n => autonumberLabel(scheme, n)).join(' ');

/** The hanging indent PowerPoint gives a text box's paragraph when bullets or numbers are turned on (EMU). */
export const HANGING = { char: 285750, autonum: 342900 } as const;
const isHanging = (marL: number, indent: number) => (marL === HANGING.char || marL === HANGING.autonum) && indent === -marL;

/**
 * The paragraph change for a bullet choice. Turning bullets on in text with no indent (a text box) gives it the
 * hanging indent PowerPoint does, so the text does not sit on the bullet; turning them off takes that indent away
 * again. Indents that came from a placeholder or that someone set are left alone.
 */
export function bulletPatch(bullet: BulletSpec, para: { marL: number; indent: number } | undefined): ParaPatch {
  const marL = para?.marL ?? 0, indent = para?.indent ?? 0;
  if (bullet.kind === 'none') return isHanging(marL, indent) ? { bullet, marL: 0, indent: 0 } : { bullet };
  const hang = HANGING[bullet.kind];
  if ((marL === 0 && indent === 0) || (isHanging(marL, indent) && marL !== hang)) return { bullet, marL: hang, indent: -hang };
  return { bullet };
}

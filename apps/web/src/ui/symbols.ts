// apps/web/src/ui/symbols.ts  -- the characters the Symbol menu offers, in groups

/** A character to insert and what it is called (its tooltip, and what a screen reader says). */
export interface SymbolChar { char: string; name: string; }

const group = (spec: string): SymbolChar[] => spec.split('|').map(s => { const [char, name] = s.split('='); return { char: char!, name: name! }; });

/** Characters people reach for that a keyboard does not have, grouped as a picker shows them. */
export const SYMBOL_GROUPS: [string, SymbolChar[]][] = [
  ['Punctuation', group('—=Em dash|–=En dash|…=Ellipsis|•=Bullet|·=Middle dot|“=Left double quote|”=Right double quote|‘=Left single quote|’=Right single quote|«=Left guillemet|»=Right guillemet|¶=Paragraph mark|§=Section mark|†=Dagger|‡=Double dagger|¿=Inverted question mark|¡=Inverted exclamation mark')],
  ['Currency and marks', group('€=Euro|£=Pound|¥=Yen|₹=Rupee|¢=Cent|₩=Won|₽=Rouble|©=Copyright|®=Registered|™=Trademark|°=Degree|№=Numero|✓=Check mark|✗=Cross mark|★=Star|☆=Outline star')],
  ['Maths', group('±=Plus or minus|×=Multiply|÷=Divide|≠=Not equal|≈=Almost equal|≤=Less than or equal|≥=Greater than or equal|∞=Infinity|√=Square root|∑=Sum|∏=Product|∫=Integral|∂=Partial derivative|∆=Increment|∈=Element of|∩=Intersection|∪=Union|⊂=Subset of|∴=Therefore|¼=One quarter|½=One half|¾=Three quarters|²=Superscript two|³=Superscript three|‰=Per mille|µ=Micro')],
  ['Greek', group('α=Alpha|β=Beta|γ=Gamma|δ=Delta|ε=Epsilon|θ=Theta|λ=Lambda|μ=Mu|π=Pi|ρ=Rho|σ=Sigma|τ=Tau|φ=Phi|ω=Omega|Γ=Capital gamma|Δ=Capital delta|Θ=Capital theta|Λ=Capital lambda|Σ=Capital sigma|Φ=Capital phi|Ψ=Capital psi|Ω=Capital omega')],
  ['Arrows', group('←=Left arrow|→=Right arrow|↑=Up arrow|↓=Down arrow|↔=Left right arrow|↕=Up down arrow|⇐=Left double arrow|⇒=Right double arrow|⇔=Left right double arrow|↵=Return arrow')],
];

/** Every character offered, in the order shown. */
export const ALL_SYMBOLS: SymbolChar[] = SYMBOL_GROUPS.flatMap(([, chars]) => chars);

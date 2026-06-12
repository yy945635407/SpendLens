# SpendLens Design System v3.0 — makemepulse Edition

## Design Philosophy

Inspired by [makemepulse](https://2019.makemepulse.com)'s immersive, cinematic web experiences:
- **Each moment gets full focus** — content breathes in generous negative space
- **Organic, fluid motion** — nothing mechanical, everything alive
- **Subtle atmosphere** — gradient backgrounds with floating particles
- **Minimal chrome** — UI recedes until needed, letting data shine
- **Narrative flow** — content reveals progressively through scroll/gesture

---

## Theme Selection

Users can switch between two themes. Both share the same structure, tokens, and animation rules — only the color palette changes.

### Theme A: Sakura Pink (default)

Warm, inviting, sunset-inspired. For users who prefer a soft, emotional connection to their finances.

```
┌──────────────────────┬──────────┬───────────────────────────────────────┐
│ Token                │ Hex      │ Role                                  │
├──────────────────────┼──────────┼───────────────────────────────────────┤
│ --theme-primary      │ #FF6B8A  │ CTAs, active states, accents          │
│ --theme-primary-light│ #FFB3C6  │ Hover states, subtle highlights       │
│ --theme-glow         │ #FFD4DF  │ Glow effects, particle color          │
│ --theme-surface      │ #FFF0F5  │ Card backgrounds, sections            │
│ --theme-gradient-1   │ #FFF0F5  │ Gradient start (light warm pink)      │
│ --theme-gradient-2   │ #FFE4EC  │ Gradient mid (soft pink)              │
│ --theme-gradient-3   │ #FFF5F8  │ Gradient end (near white)             │
│ --theme-dark         │ #5C2D3E  │ Dark surfaces (cover, summary)        │
│ --theme-positive     │ #7BC8A4  │ Income, under-budget, success         │
│ --theme-warning      │ #FFD4B8  │ Near-budget, gold accent              │
│ --theme-danger       │ #FF3D6A  │ Over-budget alert                     │
└──────────────────────┴──────────┴───────────────────────────────────────┘
```

### Theme B: Ocean Blue

Cool, calm, focused. For users who prefer a analytical, serene connection to their finances.

```
┌──────────────────────┬──────────┬───────────────────────────────────────┐
│ Token                │ Hex      │ Role                                  │
├──────────────────────┼──────────┼───────────────────────────────────────┤
│ --theme-primary      │ #5B8DEF  │ CTAs, active states, accents          │
│ --theme-primary-light│ #A8C8FF  │ Hover states, subtle highlights       │
│ --theme-glow         │ #C8DDFF  │ Glow effects, particle color          │
│ --theme-surface      │ #F0F5FF  │ Card backgrounds, sections            │
│ --theme-gradient-1   │ #F0F5FF  │ Gradient start (light cool blue)      │
│ --theme-gradient-2   │ #E4ECFF  │ Gradient mid (soft blue)              │
│ --theme-gradient-3   │ #F5F8FF  │ Gradient end (near white)             │
│ --theme-dark         │ #1E2A4A  │ Dark surfaces (cover, summary)        │
│ --theme-positive     │ #5BC8A4  │ Income, under-budget, success         │
│ --theme-warning      │ #B8D4FF  │ Near-budget, cool accent              │
│ --theme-danger       │ #EF5B8A  │ Over-budget alert (keeps pink)        │
└──────────────────────┴──────────┴───────────────────────────────────────┘
```

---

## Shared Design Tokens (theme-agnostic)

### Typography

```
┌──────────┬──────────────────────────────────────────────────┐
│ Role     │ Value                                            │
├──────────┼──────────────────────────────────────────────────┤
│ Font     │ 'Inter', -apple-system, BlinkMacSystemFont,      │
│          │ 'PingFang SC', sans-serif                        │
│ Mono     │ 'JetBrains Mono', 'Courier New', monospace       │
│ Hero     │ 36-48px, weight 200 (extra-light), letter-spacing -0.02em │
│ Title    │ 24-28px, weight 600                              │
│ Heading  │ 18-20px, weight 600                              │
│ Body     │ 15px, weight 300, line-height 1.6                │
│ Caption  │ 13px, weight 400, muted color                    │
│ Micro    │ 11px, weight 400, muted color                    │
│ Mono     │ 14px, weight 400, tabular-nums                   │
└──────────┴──────────────────────────────────────────────────┘
```

### Spacing (breathing room is key)

```
┌──────┬──────┬─────────────────────────────────────────────┐
│ Step │ Size │ Usage                                        │
├──────┼──────┼─────────────────────────────────────────────┤
│ 2xs  │ 4px  │ Tight icon-text gaps                         │
│ xs   │ 8px  │ Chip gaps, inline separators                 │
│ sm   │ 16px │ Card internal padding                        │
│ md   │ 24px │ Between related sections                     │
│ lg   │ 32px │ Between major content blocks                 │
│ xl   │ 48px │ Page-level section separation                │
│ 2xl  │ 64px │ Hero section top/bottom breathing            │
│ 3xl  │ 96px │ Full-viewport margin (desktop only)          │
└──────┴──────┴─────────────────────────────────────────────┘
```

### Border Radius

```
┌───────┬───────┬──────────────────────────────────────────┐
│ Token │ Value │ Usage                                     │
├───────┼───────┼──────────────────────────────────────────┤
│ sm    │ 8px   │ Inputs, tags, small badges                │
│ md    │ 16px  │ Cards, panels, chart containers           │
│ lg    │ 24px  │ Large cards, modals                       │
│ xl    │ 32px  │ Hero cards, feature panels                │
│ pill  │ 9999px│ Buttons, tab pills, selection indicators │
└───────┴───────┴──────────────────────────────────────────┘
```

### Shadows (soft, atmospheric — not harsh)

```
┌──────────┬───────────────────────────────────────────────────┐
│ Token    │ Value                                             │
├──────────┼───────────────────────────────────────────────────┤
│ card     │ 0 2px 20px rgba(theme-dark, 0.06)                │
│ elevated │ 0 8px 40px rgba(theme-dark, 0.08)                │
│ button   │ 0 4px 20px rgba(theme-primary, 0.25)             │
│ modal    │ 0 20px 80px rgba(theme-dark, 0.12)               │
│ glow     │ 0 0 60px rgba(theme-primary, 0.15)               │
└──────────┴───────────────────────────────────────────────────┘
```

### Motion (makemepulse-inspired organic curves)

```
┌──────────┬──────────────────────────────────────────────────┐
│ Property │ Value                                             │
├──────────┼──────────────────────────────────────────────────┤
│ Easing   │ cubic-bezier(0.22, 1, 0.36, 1) — spring-like     │
│ Fast     │ 200ms (hover, micro-interactions)                 │
│ Normal   │ 400ms (card enter, state transition)              │
│ Slow     │ 800ms (page enter, hero animation)                │
│ Stagger  │ 60-80ms per card (cascade effect)                 │
└──────────┴──────────────────────────────────────────────────┘
```

### Keyframes

```css
/* Cards rise from below with slight fade — like landscape emerging */
@keyframes riseIn {
  from { opacity: 0; transform: translateY(40px) scale(0.98); }
  to   { opacity: 1; transform: translateY(0) scale(1); }
}

/* Hero text fades in softly */
@keyframes breatheIn {
  from { opacity: 0; transform: scale(1.04); filter: blur(8px); }
  to   { opacity: 1; transform: scale(1); filter: blur(0); }
}

/* Subtle floating particles */
@keyframes float {
  0%, 100% { transform: translateY(0) translateX(0); opacity: 0.4; }
  25%      { transform: translateY(-20px) translateX(10px); opacity: 0.7; }
  50%      { transform: translateY(-10px) translateX(-5px); opacity: 0.5; }
  75%      { transform: translateY(-30px) translateX(15px); opacity: 0.3; }
}
```

---

## Component Specs

### Background Canvas

Every page has a full-viewport gradient background with subtle floating particles.

- **Pink**: `linear-gradient(135deg, #FFF0F5 0%, #FFE4EC 40%, #FFF5F8 100%)`
- **Blue**: `linear-gradient(135deg, #F0F5FF 0%, #E4ECFF 40%, #F5F8FF 100%)`
- Particles: 15-30 small circles, 2-6px diameter, theme-glow color, random float animation (8-20s duration), scattered across the viewport at z-index 0
- Optional: subtle noise/grain overlay at opacity 0.03

### Card

```
Background:    rgba(255,255,255,0.65) — frosted glass
Backdrop:      blur(12px) — iOS-style blur
Border:        1px solid rgba(theme-primary, 0.08) — subtle
Shadow:        0 2px 20px rgba(theme-dark, 0.06)
Border-radius: 16px
Animation:     riseIn 0.6s cubic-bezier(0.22,1,0.36,1), stagger by card index
```

### Primary Button

```
Background:    theme-primary
Text:          white, weight 600
Border-radius: pill
Shadow:        0 4px 20px rgba(theme-primary, 0.25)
Hover:         transform: translateY(-1px); box-shadow deepens
Active:        transform: translateY(0); shadow recedes
Transition:    200ms cubic-bezier(0.22,1,0.36,1)
```

### Tab Pill Bar

```
Container:     rgba(255,255,255,0.5), backdrop-blur(8px)
               border-radius: pill, padding: 4px
Active tab:    theme-primary, white text, pill shape
Inactive tab:  transparent, muted text
Transition:    background-color 300ms
```

### Upload Zone

```
Background:    rgba(255,255,255,0.4), dashed border (theme-glow)
Hover/Drag:    background intensifies, border solidifies, subtle scale 1.01
Empty:          dashed border visible
File selected: border solid, checkmark, file info
```

### Stat Card Grid

```
Layout:        2-3 column grid, 24px gap
Card:           glass card, left accent bar (3px, theme-primary)
Value:          28px, weight 200 (extra-light), theme-dark
Label:          13px, weight 500, theme-primary
Animation:      riseIn with 80ms stagger per card
```

### Chart Container

```
Background:    rgba(255,255,255,0.5), borderRadius: 16px
Padding:        24px
Chart:          PNG image, transparent bg, width-fill
Animation:      riseIn 0.8s
```

### Progress Indicator

```
Track:          rgba(theme-primary, 0.1), 4px height
Fill:           theme-primary, animated width transition
Label:          mono font, theme-dark
```

### Toast / Notification

```
Background:    theme-dark (solid, no blur)
Text:          white
Border-radius: 12px
Shadow:        0 8px 40px rgba(theme-dark, 0.2)
Enter:         slide up + fade, 400ms
Exit:          fade out, 200ms
```

### Modal Overlay

```
Backdrop:      rgba(255,255,255,0.8), backdrop-blur(8px)
Card:          white glass, borderRadius: 24px, shadow-modal
Enter:         scale(0.96)→scale(1) + fade, 400ms spring
```

---

## Theme Switcher

A small floating toggle in the header area:
- Two small circles (pink / blue) side by side
- Active: filled with theme-primary, slight glow
- Inactive: outline only, muted
- Transition: color + glow swap, 600ms ease

---

## Chart Palette (unified across both themes)

Charts use the active theme's primary color scale, not fixed colors:

```python
# Pink theme chart colors
CHART_PINK = ['#FF6B8A','#FF85A2','#FF9EBB','#FFB3C6','#FFD4B8','#FF7EB3','#C4909E','#7BC8A4','#FFD4DF','#E8A0B4']

# Blue theme chart colors
CHART_BLUE = ['#5B8DEF','#7BA3F5','#9BB9FB','#BBCFFF','#B8D4FF','#8BABF5','#90A4C4','#5BC8A4','#C8DDFF','#A0B4E8']
```

Both use `#5C2D3E` (pink dark) / `#1E2A4A` (blue dark) for text, and `#C4909E` (pink muted) / `#90A4C4` (blue muted) for secondary.

---

## Implementation Plan

1. **DESIGN_SYSTEM.md** — final spec (this file)
2. **templates/index.html** — web app with theme switcher, particles canvas, glass cards
3. **miniprogram/app.wxss** — CSS variables for both themes, theme class toggle
4. **ppt_builder.py** — C dict parameterized by theme, `build_ppt(data, charts, theme='pink')`
5. **pdf_builder.py** — C dict parameterized by theme
6. **charts.py** — `CHART_COLORS[theme]` dict, theme-aware palette selection

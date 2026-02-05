# Org Chart Evaluation & Improvement Plan

**Role:** Data visualization & UI/UX perspective  
**Scope:** Current org chart (jQuery OrgChart + Flask), from data layer to front-end.

---

## 1. Current State Evaluation

### 1.1 Readability (High impact)

| Issue | Where | Cause |
|-------|--------|--------|
| **Group header truncation** | "TRAINEES & EARLY CAREER" → "TRAINEES & EARLY CAREE" | Fixed `max-width: 340px` on `.group-node` + compact layout constrains the node; text overflows and is cut. |
| **Role/title truncation** | "Chief Human Resources Offi...", "Head of HR Planning & Strateg..." | Person nodes use `shortTitle` capped at 40 chars in `app.py` (37 + "…") and fixed card width (200px) with `white-space: nowrap` + `text-overflow: ellipsis`. |
| **Name truncation** | "Alhash, Reem Mohammad..." | Same fixed width; long names get ellipsis. Full text only in tooltip. |

**Verdict:** Truncation is the main barrier to quick scanning and trust in the chart. Group labels and key titles should not be cut off; person cards can use ellipsis if tooltips are obvious.

---

### 1.2 Visual Hierarchy & Consistency (Medium impact)

| Observation | Assessment |
|-------------|------------|
| **CHRO** (dark purple) vs **Group headers** (slate vs purple) | Intentional and clear: one top role, then categories. Good. |
| **Connector lines** | Template sets `.orgchart .lines .topLine, ...` to neutral gray, but the plugin also draws lines via `.hierarchy::before` / `::after` with `rgba(217,83,79,.8)` (red). Those pseudo-elements are not overridden, so red lines still appear. |
| **Leader vs staff cards** | Blue tint for leaders vs white for staff is clear. |
| **Trainee section** | After fixes, header is dark purple and readable; trainee cards are distinct. |

**Verdict:** Hierarchy is understandable. Line color inconsistency (red vs design tokens) and any remaining contrast issues should be fixed so the chart feels cohesive.

---

### 1.3 Information Architecture & Use of Space (Medium impact)

| Observation | Assessment |
|-------------|------------|
| **Three sibling groups** | LEADERSHIP & HEADS, PROFESSIONAL STAFF, TRAINEES & EARLY CAREER are siblings under CHRO. Layout is horizontal; on small viewports or when expanded, space can get tight. |
| **Placeholder nodes** | "(2) Managers" etc. are styled as aggregates (dashed, muted). Clear. |
| **Search** | Search-by-name and focus/scroll work; no visible "back to top" or "reset view". |
| **Empty space** | Plugin controls layout; large empty areas are normal for tree layouts. Zoom/pan help. |

**Verdict:** Structure is sound. Improvements: ensure group headers never truncate, consider a "Reset view" control, and optionally a small legend or "How to read" hint.

---

### 1.4 Interaction & Accessibility (Lower impact, high value)

| Observation | Assessment |
|-------------|------------|
| **Tooltips** | Full name + title + org on hover; good for truncated text. |
| **Click to focus** | Focus ring and scroll-to-node improve navigation. |
| **Keyboard / screen readers** | No evidence of focus management or ARIA; chart is likely mouse/touch-only and not fully accessible. |
| **Mobile** | Pan/zoom help, but fixed card sizes and dense branches may be hard on small screens. |

**Verdict:** Core interaction is good. Truncation reduces reliance on hover; fixing truncation improves usability. Accessibility and mobile can be Phase 2.

---

## 2. Improvement Plan (Prioritized)

### Phase 1: Readability & Critical Fixes (Do first)

**1.1 Eliminate group-header truncation**

- **Back-end:** No change needed; group names are already full ("TRAINEES & EARLY CAREER").
- **Front-end:**
  - Allow group nodes to grow by content: e.g. `min-width: 200px; max-width: none` (or a large max, e.g. 480px) for `.group-node` and `.group-node-trainees`.
  - Ensure group header text wraps (already have `white-space: normal; word-break: break-word`).
  - If the plugin forces a fixed width on compact nodes, override it for group nodes only (e.g. `.group-node.compact { width: auto; min-width: 240px; max-width: 480px; }` or equivalent so "TRAINEES & EARLY CAREER" fits on one or two lines).

**1.2 Reduce person-card truncation (without breaking layout)**

- **Back-end:** Option A: Increase `shortTitle` cap from 40 to 50–55 characters for display; keep full `title` for tooltip. Option B: Do not truncate in data; rely on CSS ellipsis and tooltip.
- **Front-end:** Slightly wider person cards (e.g. 220px) for leader roles so "Director of National Talent Management" and similar titles show more characters before ellipsis. Keep tooltip as primary source of full text.

**1.3 Unify connector line color**

- **Front-end:** Override the plugin’s red lines. The library uses `.hierarchy::before`, `.hierarchy::after`, and `.node::before`/`::after` for connectors. Add CSS that targets those pseudo-elements (e.g. `border-color: var(--org-line)` or `#cbd5e1`) so all lines match the design system (neutral gray, not red).

---

### Phase 2: Polish & Consistency

**2.1 CHRO and key titles**

- Ensure CHRO node has enough width so "Chief Human Resources Officer (CHRO)" doesn’t truncate (e.g. ensure min-width and allow wrapping for title line only if needed).
- Optional: Use `title` (full) for CHRO node content instead of `shortTitle` so the top box never truncates the C-level title.

**2.2 Tooltip affordance**

- Add a small visual cue (e.g. subtle "…" or icon) on nodes that are truncated, so users know to hover for full text.

**2.3 "Reset view" and hint**

- Add a "Reset view" or "Center chart" button that re-centers and optionally focuses the root.
- Keep or refine the existing hint ("Click a group or leader to expand their team") so it’s visible but not noisy.

---

### Phase 3: Optional Enhancements

**3.1 Responsive behavior**

- Consider slightly smaller card widths or font size at a breakpoint (e.g. max-width: 768px) so the chart remains usable on tablets.

**3.2 Export / print**

- If stakeholders need it, add "Print" or "Export to PDF" that opens a print-friendly view or triggers browser print with appropriate styles.

**3.3 Accessibility**

- Add `aria-label` to chart container and key nodes.
- Ensure focus is trapable and that keyboard users can move between nodes or at least reach search and reset.

**3.4 Legend**

- Optional one-line legend: e.g. "Purple = C-level / Trainees group; Blue border = Leaders; Gray = Staff."

---

## 3. Implementation Checklist (Phase 1)

- [ ] **Group headers:** Adjust `.group-node` (and `.group-node-trainees`) so `max-width` is larger or `none`, and override compact width for group nodes so "TRAINEES & EARLY CAREER" and "PROFESSIONAL STAFF" never truncate.
- [ ] **Person cards:** Increase person node width (e.g. to 220px) and/or increase `shortTitle` length cap in `app.py` to 50–55 chars.
- [ ] **Connector lines:** Add overrides for `.orgchart .hierarchy::before`, `::after`, and `.node::before`/`::after` so `border-color` (and if needed `background-color`) use `var(--org-line)`.
- [ ] **CHRO:** Ensure CHRO node can show full title (wider or use full title for this node only).
- [ ] **Smoke test:** Load chart, confirm no "CAREE" or similar truncation on group headers, and that lines are neutral, not red.

---

## 4. Success Criteria

- No group header text truncated (e.g. full "TRAINEES & EARLY CAREER", "PROFESSIONAL STAFF", "LEADERSHIP & HEADS").
- Connector lines use design-system gray, not red.
- CHRO title readable without truncation (or one clear ellipsis + tooltip).
- Person cards show as much title as fits at 220px (or chosen width), with tooltip for full text.
- Chart still loads and performs well with large datasets (no regression from width/wrap changes).

This plan prioritizes readability and visual consistency first, then polish and optional features.

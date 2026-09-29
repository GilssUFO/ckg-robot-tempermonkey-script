# ASIK & SurveyJS Automation Guidelines

## 1. SurveyJS & Vue Event Dispatching
- Always dispatch full synthetic mouse events (`mouseover`, `pointerdown`, `mousedown`, `pointerup`, `mouseup`, `click`) for SurveyJS elements (`.sd-question`, `label`, `.sv-string-viewer`).
- Handle dropdowns by clicking the dropdown wrapper, waiting for `.sv-popup` / `[role="listbox"]` to render, and dispatching events on target `.sv-list__item` / `[role="option"]`.
- For text and numeric inputs, override the value property setter via `HTMLInputElement.prototype` / `HTMLTextAreaElement.prototype` to ensure Vue 3 reactivity catches the updates.

## 2. Clinical Question Defaults & Safe Answers
- **General Screening**: Default to `"normal"`, `"tidak ada"`, `"tidak pernah"`, or `"tidak"`.
- **Immunization**: Default to `"ya"` or `"sudah"`.
- **Dental**: Default to `"tidak ada"` or `"0"` for caries.
- **Vision & Hearing**: Default to `"normal (visus 6/6 - 6/9)"`, `"tidak ada serumen impaksi"`, `"tidak ada infeksi"`.
- **Anthropometry (TB/BB)**: Extract values from Excel; fallback to normalized age-appropriate values (e.g., TB 110–120 cm, BB 20–25 kg for elementary school students).
- **Blood Pressure**: Extract Sistole/Diastole from Excel; fallback to 120/80 mmHg.

## 3. Navigation & State Persistence
- In ASIK Single Page Applications (SPA), use browser history navigation (`window.history.back()`) to preserve pre-selected dropdown filters (such as School and Class selections) without resetting user state.
- Use `localStorage` to maintain robot running phase and patient index across page reloads and cross-domain form navigations (`form.kemkes.go.id` <-> `sehatindonesiaku.kemkes.go.id`).
- After completing all services and Tatalaksana for a student, navigate back **2 times** (using the back arrow icon or `window.history.back()` with 1.5s delay) to return to the active class student search list.

## 4. Service-by-Service Execution Flow (Pelayanan oleh Nakes)
- On the student detail page, iterate over all table rows containing `Input Data`, `Dalam Pemeriksaan`, or `Sedang Pemeriksaan`.
- For each unfinished service, click the action button to open the form, execute SurveyJS form filling, save, and handle confirmation popups.
- Once all service rows are marked complete, proceed to `Mulai Tatalaksana`.
- Complete Tatalaksana, save, and execute 2x back navigation to continue to the next student in the queue.

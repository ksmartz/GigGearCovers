# Copilot Repo Instructions (VB.NET — Visual Studio)
_Last updated: <keep this line and update when you change rules>_

These instructions guide GitHub Copilot **Chat** when working in this repository (Visual Studio). They define how answers and code should be produced, with a strong bias toward **ready-to-compile VB.NET**.

---

## Languages & Environment
- **Primary language:** VB.NET on .NET (current LTS).
- **Editor/IDE:** Visual Studio (Windows).
- **Target:** Code must compile when pasted into Visual Studio.
- Assume `Option Strict On`, `Option Explicit On`, and `Option Infer On`.

---

## Code Response Rules (Very Important)
When Copilot Chat provides VB.NET code for this repo, it must follow **all** of these:

1) **Always return complete, runnable code**
   - Include **all required `Imports`**.
   - If the code is a `Sub` or `Function`, **return the entire Sub/Function** in a single block, **not snippets**.

2) **Update/Repair Workflow**
   - If fixing code, present **(a)** the *incorrect snippet*, **(b)** the *corrected snippet*, then **(c)** the **complete, fully updated Sub/Function**.
   - Prepend a short **header comment** above the Sub/Function including:
     - Purpose
     - Dependencies (`Imports`, NuGet packages)
     - **Current date** (use today’s date)
   - Add **clear inline comments** explaining what each section does.

3) **Sample Usage**
   - When appropriate, provide a **small “sample usage” block** (e.g., a `Sub Main()` or a short call site) so code can be tested immediately.

4) **No placeholders**
   - Avoid `...`, pseudo-code, or “left to the reader.” Use realistic names and minimal but **real** example data.

5) **Error Handling**
   - Use `Try...Catch` with **meaningful messages**. Don’t swallow exceptions silently.
   - Prefer `Using` blocks for disposables (`HttpClient`, `SqlConnection`, file streams).

6) **Async patterns**
   - Prefer `Async`/`Await` where appropriate; do **not** block on tasks with `.Result` or `.Wait()` unless there’s a good reason.

7) **Formatting & Style**
   - **Naming:** PascalCase for classes/modules/methods; camelCase for locals/parameters; ALL_CAPS for constants.
   - Be explicit with types (no ambiguous variants).
   - Keep lines readable; split long expressions.

---

## Typical Tasks (Give solutions in this style)
- **HTTP / Web APIs:** Use `HttpClient` in a `Using` block; set headers; serialize with `Newtonsoft.Json` if needed. Show a clean model type if you parse JSON.
- **SQL Server (Developer/Express 2022):** Use `System.Data.SqlClient` / `Microsoft.Data.SqlClient` with parameterized commands. Never concatenate SQL.
- **WinForms/WPF utilities:** Provide a complete form/window (e.g., 800×800) with event handlers wired up. Explain where to place the code (Form class vs. Module).
- **File IO:** Use proper encoding; validate existence; wrap in `Try...Catch`.

---

## Repository Context & Preferences
- Assume most deliverables are **tools/utilities** that will be **pasted directly** into a project and run.
- If multiple alternatives exist, prefer the approach that is **simplest to compile** in a clean Visual Studio install.
- If an external package is required, **list the exact NuGet package name** and add a short install note.

---

## Example Answer Pattern (Template)
> Use this **structure** whenever fixing or updating a Sub/Function.

**Incorrect snippet**
```vbnet
' (Show the minimal faulty code that was provided or assumed)

"""Standardized 담당팀 search suggestions with non-standard confirmation."""
import tkinter as tk
from tkinter import ttk

import main_enterprise_v3 as v3
import ui_enterprise as ui

TEAMS = (
    "파우치형Pack개발품질1팀",
    "파우치형Pack개발품질2팀",
    "원통형Pack개발품질팀",
    "ESS System개발품질팀",
    "회로개발품질팀",
    "양산이슈",
)

_original_entry = v3.EnterpriseAppV3._enterprise_entry
_original_run = v3.EnterpriseAppV3.run


def _team_autocomplete_entry(self, parent, label, key, hint):
    if key != "team":
        return _original_entry(self, parent, label, key, hint)

    row = tk.Frame(parent, bg="white")
    row.pack(fill="x", pady=4)
    tk.Label(row, text=label, bg="white", fg=ui.TEXT,
             font=("Malgun Gothic", 9), width=17, anchor="w").pack(side="left")

    entry_wrap = tk.Frame(row, bg="white")
    entry_wrap.pack(side="left", fill="x", expand=True)
    entry = ttk.Entry(entry_wrap, textvariable=self.vars[key])
    entry.pack(fill="x")

    popup = None
    listbox = None

    def matches_for(value):
        q = value.strip().casefold()
        return [x for x in TEAMS if q in x.casefold()] if q else list(TEAMS)

    def close_popup(_event=None):
        nonlocal popup, listbox
        if popup is not None:
            try:
                popup.destroy()
            except Exception:
                pass
        popup = None
        listbox = None

    def choose_candidate(_event=None):
        if listbox is None:
            return
        sel = listbox.curselection()
        if not sel:
            return
        self.vars[key].set(listbox.get(sel[0]))
        entry.icursor("end")
        close_popup()
        entry.focus_set()

    def show_candidates(_event=None):
        nonlocal popup, listbox
        value = self.vars[key].get()
        candidates = matches_for(value)
        close_popup()
        if not candidates:
            return
        self.update_idletasks()
        popup = tk.Toplevel(self)
        popup.wm_overrideredirect(True)
        popup.configure(bg=ui.BORDER)
        try:
            popup.attributes("-topmost", True)
        except Exception:
            pass
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        width = max(entry.winfo_width(), 300)
        height = min(4, len(candidates)) * 28 + 4
        popup.geometry(f"{width}x{height}+{x}+{y}")
        listbox = tk.Listbox(
            popup, font=("Malgun Gothic", 9), relief="flat", bd=0,
            highlightthickness=1, highlightbackground=ui.BORDER,
            activestyle="none", exportselection=False
        )
        listbox.pack(fill="both", expand=True, padx=1, pady=1)
        for item in candidates:
            listbox.insert("end", item)
        listbox.bind("<ButtonRelease-1>", choose_candidate)
        listbox.bind("<Return>", choose_candidate)
        listbox.bind("<Escape>", close_popup)

    def on_keyrelease(event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        show_candidates()

    def on_down(_event=None):
        if popup is None:
            show_candidates()
        if listbox is not None and listbox.size():
            listbox.focus_set()
            listbox.selection_clear(0, "end")
            listbox.selection_set(0)
            listbox.activate(0)
        return "break"

    def delayed_close(_event=None):
        # Give a mouse click on the suggestion list time to complete first.
        self.after(180, lambda: close_popup() if entry.focus_get() is not listbox else None)

    entry.bind("<KeyRelease>", on_keyrelease, add="+")
    entry.bind("<FocusIn>", show_candidates, add="+")
    entry.bind("<Down>", on_down, add="+")
    entry.bind("<Escape>", close_popup, add="+")
    entry.bind("<FocusOut>", delayed_close, add="+")

    tk.Label(parent, text="표준 팀명 검색 후보 표시 · 클릭 시 입력 · 비표준 팀명은 실행 전 확인",
             bg="white", fg="#98A3AD", font=("Malgun Gothic", 7)).pack(anchor="e")


def _run_with_team_confirmation(self):
    team = self.vars.get("team").get().strip() if self.vars.get("team") else ""
    if team and team not in TEAMS:
        message = (
            f"입력한 담당팀 '{team}'은 표준 팀명 목록에 해당하지 않습니다.\n\n"
            "그럼에도 이 팀명으로 업데이트를 실행하시겠습니까?"
        )
        if not ui.ask_yes_no(self, "담당팀 확인", message):
            try:
                self.status_var.set("READY · 담당팀을 다시 확인해 주세요.")
            except Exception:
                pass
            return
    return _original_run(self)


v3.EnterpriseAppV3._enterprise_entry = _team_autocomplete_entry
v3.EnterpriseAppV3.run = _run_with_team_confirmation

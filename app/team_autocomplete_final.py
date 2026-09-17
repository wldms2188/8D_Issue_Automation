"""Standardized 담당팀 autocomplete for the enterprise UI."""
import tkinter as tk
from tkinter import ttk

import main_enterprise_v3 as v3
import ui_enterprise as ui

TEAMS = (
    "파우치형Pack개발품질1팀",
    "파우치형Pack개발품질2팀",
    "원통형Pack개발품질팀",
    "ESS System개발품질팀",
)

_original_entry = v3.EnterpriseAppV3._enterprise_entry


def _team_autocomplete_entry(self, parent, label, key, hint):
    if key != "team":
        return _original_entry(self, parent, label, key, hint)

    row = tk.Frame(parent, bg="white")
    row.pack(fill="x", pady=4)
    tk.Label(
        row, text=label, bg="white", fg=ui.TEXT,
        font=("Malgun Gothic", 9), width=17, anchor="w"
    ).pack(side="left")

    combo = ttk.Combobox(
        row, textvariable=self.vars[key], values=TEAMS, state="normal"
    )
    combo.pack(side="left", fill="x", expand=True)

    def filter_values(_event=None):
        typed = self.vars[key].get().strip().casefold()
        matches = [x for x in TEAMS if typed in x.casefold()] if typed else list(TEAMS)
        combo.configure(values=matches or TEAMS)

    def complete_selection(_event=None):
        value = self.vars[key].get().strip()
        if value in TEAMS:
            return
        matches = [x for x in TEAMS if value.casefold() in x.casefold()]
        if len(matches) == 1:
            self.vars[key].set(matches[0])

    combo.bind("<KeyRelease>", filter_values, add="+")
    combo.bind("<<ComboboxSelected>>", complete_selection, add="+")
    combo.bind("<FocusOut>", complete_selection, add="+")

    tk.Label(
        parent,
        text="표준 팀명 선택 · 입력 시 자동완성",
        bg="white", fg="#98A3AD", font=("Malgun Gothic", 7)
    ).pack(anchor="e")


v3.EnterpriseAppV3._enterprise_entry = _team_autocomplete_entry

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
    tk.Label(row, text=label, bg="white", fg=ui.TEXT,
             font=("Malgun Gothic", 9), width=17, anchor="w").pack(side="left")

    combo = ttk.Combobox(row, textvariable=self.vars[key], values=TEAMS, state="normal")
    combo.pack(side="left", fill="x", expand=True)

    def matches_for(value):
        q = value.strip().casefold()
        return [x for x in TEAMS if q in x.casefold()] if q else list(TEAMS)

    def filter_values(_event=None):
        combo.configure(values=matches_for(self.vars[key].get()) or TEAMS)

    def commit_standard(_event=None):
        value = self.vars[key].get().strip()
        if value in TEAMS:
            combo.configure(values=TEAMS)
            return
        matches = matches_for(value)
        if len(matches) == 1:
            self.vars[key].set(matches[0])
        else:
            # Never leave a non-standard team name in the business data.
            self.vars[key].set("")
        combo.configure(values=TEAMS)

    combo.bind("<KeyRelease>", filter_values, add="+")
    combo.bind("<<ComboboxSelected>>", commit_standard, add="+")
    combo.bind("<Return>", commit_standard, add="+")
    combo.bind("<FocusOut>", commit_standard, add="+")

    tk.Label(parent, text="표준 팀명 4개 · 입력 시 자동완성 · 비표준 값은 저장되지 않음",
             bg="white", fg="#98A3AD", font=("Malgun Gothic", 7)).pack(anchor="e")


v3.EnterpriseAppV3._enterprise_entry = _team_autocomplete_entry

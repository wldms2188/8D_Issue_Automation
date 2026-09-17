"""Standardized 담당팀 autocomplete with non-standard confirmation."""
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
_original_run = v3.EnterpriseAppV3.run


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

    def complete_if_unique(_event=None):
        value = self.vars[key].get().strip()
        if value in TEAMS:
            combo.configure(values=TEAMS)
            return
        matches = matches_for(value)
        if len(matches) == 1:
            self.vars[key].set(matches[0])
        # Multiple/no matches: keep the user's text. Validation happens at Run.
        combo.configure(values=TEAMS)

    combo.bind("<KeyRelease>", filter_values, add="+")
    combo.bind("<<ComboboxSelected>>", complete_if_unique, add="+")
    combo.bind("<Return>", complete_if_unique, add="+")
    combo.bind("<FocusOut>", complete_if_unique, add="+")

    tk.Label(parent, text="표준 팀명 4개 · 입력 시 자동완성 · 비표준 팀명은 실행 전 확인",
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

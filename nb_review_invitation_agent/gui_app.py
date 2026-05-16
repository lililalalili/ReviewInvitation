from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from .config import RuntimeConfig
from .gui_controller import (
    DISPLAY_FIELDS,
    WorkbookReviewController,
    build_email_search_url,
    default_open_url,
)
from .invitation_service import InvitationService
from .mailer_outlook import OutlookMailer
from .template_renderer import TemplateRenderer


class ReviewGuiApp:
    def __init__(self, root: tk.Tk, controller: WorkbookReviewController, config: RuntimeConfig) -> None:
        self.root = root
        self.controller = controller
        self.config = config
        self.root.title("NB Review Invitation")
        self._editable_vars: dict[str, tk.StringVar] = {}
        self._readonly_vars: dict[str, tk.StringVar] = {}
        self.invitation_service = InvitationService(
            controller=self.controller,
            renderer=TemplateRenderer(),
            mailer=OutlookMailer(),
            confirm_send=lambda prompt: messagebox.askyesno("确认发送", prompt),
        )
        self._build_ui()
        self._refresh_view()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=8, pady=8)
        ttk.Label(top, text="Batch ID").pack(side="left")
        self.batch_var = tk.StringVar(value=self.controller.selected_batch)
        self.batch_box = ttk.Combobox(top, textvariable=self.batch_var, values=self.controller.batch_ids, state="readonly")
        self.batch_box.pack(side="left", padx=6)
        self.batch_box.bind("<<ComboboxSelected>>", lambda _: self._on_batch_change())

        body = ttk.Frame(self.root)
        body.pack(fill="both", expand=True, padx=8)

        for idx, field in enumerate(DISPLAY_FIELDS):
            ttk.Label(body, text=field).grid(row=idx, column=0, sticky="w", padx=4, pady=2)
            if field == "Overseas":
                var = tk.StringVar()
                self._editable_vars[field] = var
                cb = ttk.Combobox(body, textvariable=var, values=["Yes", "No"], state="readonly")
                cb.grid(row=idx, column=1, sticky="ew", padx=4, pady=2)
            elif field == "Manual Decision":
                var = tk.StringVar()
                self._editable_vars[field] = var
                cb = ttk.Combobox(body, textvariable=var, values=["Review", "Insight", "No"], state="readonly")
                cb.grid(row=idx, column=1, sticky="ew", padx=4, pady=2)
            elif field == "Research field":
                var = tk.StringVar()
                self._editable_vars[field] = var
                ttk.Entry(body, textvariable=var).grid(row=idx, column=1, sticky="ew", padx=4, pady=2)
            elif field in {"Last Author Web", "Pubmed Link", "Email of the Last Author"}:
                var = tk.StringVar()
                self._readonly_vars[field] = var
                ttk.Button(body, textvariable=var, command=lambda f=field: self._on_link(f)).grid(
                    row=idx, column=1, sticky="ew", padx=4, pady=2
                )
            else:
                var = tk.StringVar()
                self._readonly_vars[field] = var
                ttk.Entry(body, textvariable=var, state="readonly").grid(row=idx, column=1, sticky="ew", padx=4, pady=2)

        body.columnconfigure(1, weight=1)

        nav = ttk.Frame(self.root)
        nav.pack(fill="x", padx=8, pady=8)
        ttk.Button(nav, text="Previous row", command=self._prev).pack(side="left", padx=2)
        ttk.Button(nav, text="Next row", command=self._next).pack(side="left", padx=2)
        ttk.Button(nav, text="First row in current batch", command=self._first).pack(side="left", padx=2)
        ttk.Button(nav, text="Last row in current batch", command=self._last).pack(side="left", padx=2)
        ttk.Button(nav, text="Save all modifications", command=self._save).pack(side="left", padx=2)
        ttk.Button(nav, text="向当前作者邀稿", command=self._invite_current).pack(side="left", padx=2)
        ttk.Button(nav, text="BatchInvitation", command=self._invite_batch).pack(side="left", padx=2)

    def _apply_edits(self) -> None:
        self.controller.update_current_row({k: v.get() for k, v in self._editable_vars.items()})

    def _refresh_view(self) -> None:
        row = self.controller.get_current_row()
        for field, var in self._editable_vars.items():
            var.set(row.values.get(field, ""))
        for field, var in self._readonly_vars.items():
            var.set(row.values.get(field, ""))

    def _on_batch_change(self) -> None:
        self._apply_edits()
        self.controller.set_batch(self.batch_var.get())
        self._refresh_view()

    def _prev(self):
        self._apply_edits(); self.controller.navigate_previous(); self._refresh_view()

    def _next(self):
        self._apply_edits(); self.controller.navigate_next(); self._refresh_view()

    def _first(self):
        self._apply_edits(); self.controller.navigate_first(); self._refresh_view()

    def _last(self):
        self._apply_edits(); self.controller.navigate_last(); self._refresh_view()

    def _save(self):
        self._apply_edits()
        try:
            self.controller.save_workbook(self.config.target_xlsm)
        except Exception as exc:
            messagebox.showerror("Save failed", f"Failed to save workbook: {exc}")
            return
        messagebox.showinfo("Saved", "Workbook saved successfully")

    def _on_link(self, field: str) -> None:
        row = self.controller.get_current_row()
        value = row.values.get(field, "").strip()
        if not value:
            messagebox.showwarning("Missing link", f"No value for {field}")
            return
        if field == "Email of the Last Author":
            default_open_url(build_email_search_url(value))
        else:
            default_open_url(value)

    def _invite_current(self):
        self._apply_edits()
        result = self.invitation_service.invite_current(self.controller.get_current_row())
        if result.status == "Sent":
            self._refresh_view()
            messagebox.showinfo("邀请结果", result.message)
        elif result.status == "Cancelled":
            messagebox.showwarning("邀请取消", result.message)
        elif result.status == "Skipped":
            messagebox.showwarning("已跳过", result.message)
        else:
            messagebox.showerror("邀请失败", result.message)

    def _invite_batch(self):
        self._apply_edits()
        results = self.invitation_service.invite_batch(self.controller.get_selected_batch_rows())
        sent = sum(1 for r in results if r.status == "Sent")
        skipped = sum(1 for r in results if r.status == "Skipped")
        cancelled = sum(1 for r in results if r.status == "Cancelled")
        errors = [r for r in results if r.status == "Error"]
        self._refresh_view()
        summary = f"Sent={sent}, Skipped={skipped}, Cancelled={cancelled}, Error={len(errors)}"
        if errors:
            messagebox.showerror("BatchInvitation", f"{summary}\n首个错误: {errors[0].message}")
        else:
            messagebox.showinfo("BatchInvitation", summary)


def launch_gui(config: RuntimeConfig) -> None:
    controller = WorkbookReviewController.from_xlsm_path(config.target_xlsm)
    root = tk.Tk()
    ReviewGuiApp(root, controller, config)
    root.mainloop()

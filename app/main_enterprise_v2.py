"""Enterprise UI v2: input guidance and Issue DB PLM validation."""
import main_enterprise as ent
import ui_enterprise as ui


class EnterpriseAppV2(ent.EnterpriseApp):
    def __init__(self):
        super().__init__()
        self._replace_label_text('예: DUT3 / Sample No.', '예: A1, B1, C2')
        self._replace_label_text('기존값이 있을 때 입력', 'Issue DB 업데이트 시 필수 입력')

    def _replace_label_text(self, old, new):
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    if child.cget('text') == old:
                        child.configure(text=new)
                except Exception:
                    pass
                walk(child)
        walk(self)

    def run(self):
        g=self.gui()
        # PMS/PLM 번호는 Issue DB가 실제 업데이트 대상일 때만 필수.
        # 주간회의 PPT만 업데이트하는 경우에는 입력하지 않아도 된다.
        if g.get('xlsx','').strip() and not g.get('plm_no','').strip():
            self.status_var.set('READY  ·  PMS/PLM 이슈번호 입력이 필요합니다.')
            return ui.warning(
                self,
                'PMS/PLM 이슈번호 확인',
                'Issue DB Excel을 업데이트하려면 PMS/PLM 이슈번호를 입력해 주세요.\n\n'
                '주간회의 PPT만 업데이트하는 경우에는 입력하지 않아도 됩니다.'
            )
        return super().run()


if __name__ == '__main__':
    EnterpriseAppV2().mainloop()

# Recovery entrypoint for the last pre-GUI-change baseline.
# IMPORTANT: do not modify extraction / Excel / weekly rendering here.
# This file only imports v3.2.0 logic and gives the window an unambiguous title.

import main_v320 as v320

base = v320.base


class RecoveryApp(base.App):
    def __init__(self):
        super().__init__()
        self.title('8D 이슈 자동화 v3.2.0 RECOVERY')


if __name__ == '__main__':
    # main_v320 import already wires:
    #   base.extract = v3.1.8 extraction / Issue DB behavior
    #   base.weekly  = v3.2.0 weekly renderer
    RecoveryApp().mainloop()

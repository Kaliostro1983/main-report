# Git helpers

- **commit_and_push.bat** — додає всі зміни, робить коміт з авто-повідомленням `Auto commit YYYY-MM-DD_HH-mm-ss`, робить `pull --rebase --autostash` і `push`. Пише лог у `logs\TIMESTAMP_commit_and_push.log` і не закриває вікно.
- **update_repo.bat** — підтягнути останні зміни з GitHub; показує `HEAD(before) -> HEAD(after)` та останні коміти; лог у `logs\TIMESTAMP_update_repo.log`.
- **force_update.bat** — аварійне вирівнювання `reset --hard origin/main` (стирає незакомічені зміни).

Обидва основні .bat **самостійно знаходять корінь репозиторію** (йдуть вгору від місця, де лежить .bat, доки не знайдуть `.git`). Тож їх можна зберігати як у корені, так і в підпапках (наприклад, `scripts\`).

Якщо запускаєш із PowerShell — все працює без змін; ці скрипти відкривають **cmd**-вікно і лишають його відкритим через `pause`.

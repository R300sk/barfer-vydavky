Krátky “ťahák” na budúce

| Akcia                  | Príkaz                               |
| ---------------------- | ------------------------------------ |
| skontroluj stav        | `git status`                         |
| stiahni cudzie zmeny   | `git pull --rebase origin main`      |
| commitni nové zmeny    | `git add . && git commit -m "popis"` |
| pushni na GitHub       | `git push`                           |
| pushni do Apps Scriptu | `clasp push`                         |
| spusti funkciu v GAS   | `clasp run "nazovFunkcie"`           |


Chceš, aby som ti z týchto najčastejších príkazov spravil malý lokálny alias/skript (napr. ./sync.sh), ktorý by automaticky robil pull → push → clasp push jedným príkazom?

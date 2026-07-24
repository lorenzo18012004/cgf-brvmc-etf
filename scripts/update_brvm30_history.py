"""
update_brvm30_history.py — Mise à jour quotidienne de l'historique BRVMC (indice composite)
Source principale : sikafinance.com/marches/historiques/BRVMC
Usage : python update_brvm30_history.py
"""

import os, sys, json
from datetime import date

from base import BaseScript


class Brvm30HistoryUpdater(BaseScript):

    def _load(self, fname):
        path = os.path.join(self.data_dir, fname)
        if not os.path.exists(path):
            return None
        with open(path, encoding='utf-8') as f:
            return json.load(f)

    def _save(self, fname, data):
        path = os.path.join(self.data_dir, fname)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def run(self):
        sys.path.insert(0, self.scripts_dir)
        from scrape_sika import SikaScraper

        brvm = self._load('brvm30_index_history.json') or {}
        today_str = date.today().isoformat()

        # Source principale : page historiques/BRVMC sur Sika Finance
        try:
            hist_data = SikaScraper().scrape_brvmc_history()
            new_entries = 0
            for d, v in hist_data.items():
                if d not in brvm or brvm[d] != v:
                    brvm[d] = v
                    new_entries += 1
            today_val = hist_data.get(today_str)
            print(f"  BRVMC Sika : {len(hist_data)} dates, {new_entries} nouvelles"
                  + (f" — aujourd'hui : {today_val}" if today_val else ""))
        except Exception as e:
            print(f"  [WARN] Scraping BRVMC echoue : {e}")

        brvm_sorted = dict(sorted(brvm.items()))
        self._save('brvm30_index_history.json', brvm_sorted)
        print(f"brvm30_index_history.json mis a jour. Total : {len(brvm_sorted)} jours.")
        return brvm_sorted


if __name__ == '__main__':
    Brvm30HistoryUpdater().run()

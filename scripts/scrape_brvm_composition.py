"""
scrape_brvm_composition.py — Mise à jour de la composition BRVMC
================================================================
Pour le BRVM Composite, la "composition" = tous les tickers avec des prix
recents dans sika_history.json (au moins un prix dans les 30 derniers jours).
Pas de scraping web/PDF — la source de verite est sika_history.json.

Compare avec brvm_composition_latest.json pour detecter entrees/sorties.
Alimente propose_rebalancing.py en aval.

Usage :
    python scrape_brvm_composition.py
    python scrape_brvm_composition.py --force
"""

import os, sys, argparse
from datetime import datetime, date, timedelta, timezone

from base import BaseScript


class BRVMCompositionScraper(BaseScript):

    def __init__(self):
        super().__init__()
        self.out_file  = os.path.join(self.data_dir, "brvm_composition_latest.json")
        self.hist_file = os.path.join(self.data_dir, "brvm_composition_history.json")

    def _get_active_tickers(self, sika_history, days=30):
        """Tickers ayant au moins un prix dans les X derniers jours."""
        cutoff = (date.today() - timedelta(days=days)).isoformat()
        active = []
        for tk, hist in sika_history.items():
            if any(d >= cutoff for d in hist):
                active.append(tk)
        return sorted(active)

    def _nearest_quarter_start(self):
        """Premier jour ouvré du trimestre courant (jan/avr/jul/oct)."""
        today = date.today()
        quarter_month = ((today.month - 1) // 3) * 3 + 1
        q = date(today.year, quarter_month, 1)
        while q.weekday() >= 5:
            q += timedelta(days=1)
        return q.isoformat()

    def run(self, force=False):
        today = date.today().isoformat()

        sika_history = self.load_json("sika_history.json") or {}
        if not sika_history:
            print("[WARN] sika_history.json vide — rien a faire.")
            return False

        active = self._get_active_tickers(sika_history, days=30)
        print(f"[INFO] Composition BRVMC : {len(active)} tickers actifs (30 derniers jours)")

        current     = self.load_json("brvm_composition_latest.json") or {}
        current_set = set(current.get("composition", []))
        new_set     = set(active)

        if not force and new_set == current_set:
            print(f"[INFO] Composition inchangee. Rien a faire.")
            return False

        entries = sorted(new_set - current_set)
        exits   = sorted(current_set - new_set)

        if entries: print(f"  Entrants : {entries}")
        if exits:   print(f"  Sortants : {exits}")

        rebal_date = self._nearest_quarter_start()

        new_comp = {
            "rebal_date":  rebal_date,
            "scrape_ts":   datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "n_tickers":   len(active),
            "composition": active,
            "entries":     entries,
            "exits":       exits,
        }
        self.save_json("brvm_composition_latest.json", new_comp)
        print(f"[OK] brvm_composition_latest.json mis a jour ({rebal_date}, {len(active)} tickers).")

        hist = self.load_json("brvm_composition_history.json") or {}
        if isinstance(hist, list):
            hist = {}
        hist[today] = active
        self.save_json("brvm_composition_history.json", hist)

        return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Composition BRVMC depuis sika_history.json")
    parser.add_argument("--force", action="store_true",
                        help="Forcer la mise a jour meme si composition inchangee")
    args = parser.parse_args()
    BRVMCompositionScraper().run(force=args.force)

import fitz  # PyMuPDF

doc = fitz.open("main.pdf")
full_text = " ".join(p.get_text() for p in doc)

checks = {
    "58 wilayas"       : 0,   # doit être absent
    "69 wilayas"       : None, # doit être présent
    "97,4"             : None, # taux de réussite unifié
    "69,6"             : 0,   # doit être absent
    "147/151"          : None, # doit être présent
    "144/207"          : 0,   # doit être absent
    "??"               : 0,   # aucune référence cassée
}

for terme, attendu in checks.items():
    count = full_text.count(terme)
    statut = "✅" if (attendu is None and count > 0) \
             or count == attendu else "❌"
    print(f"{statut}  '{terme}' — {count} occurrence(s)")

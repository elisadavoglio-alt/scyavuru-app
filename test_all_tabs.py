import pandas as pd
from app import hunt_email, _apply_boolean_filters, run_search, run_competitor_posts, run_post_reactions

print("=== TEST BOOLEANI (Tab 1 e 2) ===")
# Creiamo un DataFrame fittizio
df_mock = pd.DataFrame({
    "Nome": ["Mario Rossi", "Giulia Bianchi", "Luca Verdi"],
    "Qualifica": ["Buyer GDO", "Category Manager", "Marketing Assistant"],
    "Azienda": ["Conad", "Esselunga", "Carrefour"]
})
print("DataFrame Originale:")
print(df_mock)

print("\nTest Includi: 'Buyer' O 'Manager'")
df_inc = _apply_boolean_filters(df_mock, must_include="Buyer, Manager", must_exclude="")
print(df_inc)

print("\nTest Escludi: 'Marketing'")
df_exc = _apply_boolean_filters(df_mock, must_include="", must_exclude="Marketing")
print(df_exc)

print("\n=== TEST RICERCA LIBERA (Tab 1) ===")
print("Eseguo ricerca: 'Buyer' 'Conad' 'Italia' (max 2 risultati)")
try:
    results_tab1 = run_search(ruolo="Buyer", azienda="Conad", location="Italia", max_profili=2)
    print(f"Trovati {len(results_tab1)} risultati.")
    for r in results_tab1:
        print(f" - {r.get('Nome', '')} | {r.get('Qualifica', '')} | {r.get('Email', '')}")
except Exception as e:
    print(f"Errore in Tab 1: {e}")

print("\n=== TEST POST COMPETITOR E REAZIONI (Tab 3 e 4) ===")
print("Eseguo test scraping per: 'pistì' (max 1 post)")
try:
    posts = run_competitor_posts(company_slug="pistì", max_posts=1)
    if posts:
        post = posts[0]
        print(f"Trovato Post! (Date: {post.get('date', 'N/A')})")
        print(f"Snippet: {post.get('snippet', '')[:50]}...")
        url = post.get('postUrl')
        
        print("\nTest Reazioni al post...")
        reactions = run_post_reactions(post_url=url, max_reactions=2)
        print(f"Trovate {len(reactions)} reazioni.")
        for react in reactions:
            print(f" - {react.get('Nome')} | {react.get('Qualifica')}")
    else:
        print("Nessun post trovato.")
except Exception as e:
    print(f"Errore in Tab 3/4: {e}")

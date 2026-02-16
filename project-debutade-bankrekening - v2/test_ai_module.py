"""
Test voor de tag recommender AI module
"""
import os
import sys

# Voeg project root toe aan path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from tag_recommender import TagRecommender


def test_basic_functionality():
    """Test of de recommender basis functionaliteit heeft"""
    print("Test 1: Initialisatie")
    training_path = os.path.join("static", "category_test_set.xlsx")
    tags = ["8700;Koffie", "4500;Huur gebouw", "4930;Kantoorartikelen"]
    
    recommender = TagRecommender(training_path, allowed_tags=tags)
    print(f"  ✓ Recommender aangemaakt met {len(tags)} toegestane tags")
    
    print("\nTest 2: Laden trainingsdata")
    if not os.path.exists(training_path):
        print(f"  ⚠ Trainingsbestand niet gevonden: {training_path}")
        print("  Dit is normaal als het bestand nog niet is aangemaakt.")
        return
    
    success = recommender.load()
    if success:
        print(f"  ✓ Trainingsdata geladen: {recommender.total_docs} documenten")
        print(f"  ✓ Unieke tokens: {len(recommender.token_doc_freq)}")
        print(f"  ✓ Tags in model: {len(recommender.tag_token_freq)}")
    else:
        print("  ✗ Laden mislukt")
        return
    
    print("\nTest 3: Suggestie genereren")
    test_transaction = {
        "mededelingen": "ING betaling koffie supplies voor kantoor",
        "rekening": "NL01INGB0001234567",
        "bedrag": "25.50"
    }
    
    suggestions = recommender.recommend(test_transaction, top_k=3)
    if suggestions:
        print(f"  ✓ {len(suggestions)} suggesties gegenereerd:")
        for i, sug in enumerate(suggestions, 1):
            print(f"    {i}. {sug['tag']} (score: {sug['score']})")
    else:
        print("  ⚠ Geen suggesties gevonden (mogelijk onvoldoende trainingsdata)")
    
    print("\nTest 4: Lege transactie")
    empty_transaction = {}
    suggestions = recommender.recommend(empty_transaction)
    if not suggestions:
        print("  ✓ Correct: geen suggesties voor lege transactie")
    else:
        print("  ⚠ Onverwacht: suggesties voor lege transactie")
    
    print("\nTest 5: Onbekende woorden")
    unknown_transaction = {
        "mededelingen": "xyzqweasd zxcvbnm poiuytrewq"
    }
    suggestions = recommender.recommend(unknown_transaction)
    if not suggestions:
        print("  ✓ Correct: geen suggesties voor compleet onbekende woorden")
    else:
        print(f"  ⚠ Onverwacht: {len(suggestions)} suggesties voor onbekende woorden")
    
    print("\n" + "="*60)
    print("Tests voltooid!")


if __name__ == "__main__":
    print("="*60)
    print("TAG RECOMMENDER AI MODULE - TESTS")
    print("="*60)
    print()
    
    try:
        test_basic_functionality()
    except Exception as e:
        print(f"\n✗ FOUT tijdens testen: {e}")
        import traceback
        traceback.print_exc()

"""Checks that the web lookup follows the engine rules."""
import re

import engine


def expect(condition, message):
    if not condition:
        raise SystemExit(message)


def main():
    expect(engine.LOAD_ERROR is None, engine.LOAD_ERROR)
    expect(len(engine.DTC_DATABASE) == 3717, f"expected 3717 codes, got {len(engine.DTC_DATABASE)}")

    invalid = engine.lookup_dtc("hello")
    expect(not invalid["ok"] and invalid["error"] == "invalid_format", invalid)
    expect(invalid["message"].startswith("Neteisingas formatas."), invalid["message"])

    empty = engine.lookup_dtc("   ")
    expect(empty["error"] == "invalid_format", empty)

    spaced = engine.lookup_dtc("p 0101")
    expect(spaced["ok"] and spaced["code"] == "P0101", spaced)
    direct = engine.lookup_dtc("P0101")
    expect(spaced["translation"] == direct["translation"], "spacing changed the translation")
    expect(direct["translation"] == engine.translate_description(direct["description"]), "lookup translation drifted")
    expect(len(direct["faults"]) == 3, direct["faults"])
    expect(direct["status"] == "STANDARTINIS OBD-II DIAPAZONAS", direct["status"])
    expect("ford_note" not in direct, direct)

    absent = next(
        code
        for number in range(10000)
        if (code := f"P{number:04d}") not in engine.DTC_DATABASE
    )
    missing = engine.lookup_dtc(absent)
    expect(missing["error"] == "not_found", missing)
    expect("nepateiktame šaltinyje" in missing["detail"], missing["detail"])

    for example in ("P0101", "B1202", "C1095", "U0100"):
        found = engine.lookup_dtc(example)
        expect(found["ok"], f"{example} missing from the source file: {found}")

    blower = engine.replace_dictionary_phrases("Blower Motor Low Input")
    expect("salono ventiliatorius" in blower, blower)
    expect("per žemas įėjimo signalas" in blower, blower)
    expect(engine.replace_dictionary_phrases("BlowerLow") == "BlowerLow", "Low matched inside Blower")

    lost = engine.translate_description("Lost Communication With Engine Control Module")
    expect(lost == "Prarastas ryšys su moduliu: variklio valdymo modulis.", lost)

    exact = engine.translate_description("Exhaust Camshaft Position Actuator Circuit / Open (Bank 1)")
    expect(exact.startswith("Išmetimo skirstomojo veleno padėties reguliatoriaus elektros grandinė nutrūkusi"), exact)
    expect("1-oji cilindrų eilė" in exact, exact)

    branches = {
        "Oxygen Sensor Circuit": "lambda",
        "Mass or Volume Air Flow": "oro srauto",
        "Injector Circuit": "Purkštuko",
        "Glow Plug Heater": "kaitinimo žvakė",
        "Turbocharger underboost": "Turbokompresoriaus",
        "Exhaust Gas Recirculation flow": "EGR",
        "Engine Coolant Temperature": "Aušinimo skysčio",
        "Crankshaft Position Sensor": "Alkūninio",
        "Misfire Detected": "Uždegimo sistemos",
        "Circuit Low Voltage": "Pažeisti laidai",
        "Control Module Memory": "Valdymo bloko",
        "Random unrecognized component": "Galimas atitinkamo jutiklio",
    }
    for description, fragment in branches.items():
        faults = engine.possible_faults("P0101", description)
        expect(len(faults) == 3 and fragment in faults[0], f"{description} -> {faults}")

    manufacturer = next(
        code for code in engine.DTC_DATABASE if re.fullmatch(r"P[13]\d{3}|[BCU]1\d{3}", code)
    )
    spec = engine.lookup_dtc(manufacturer)
    expect(spec["status"] == "GALIMAI GAMINTOJO SPECIFINIS KODAS", spec)

    engine.FORD_SPECIFIC_CODES["P0101"] = "Testinė Ford pastaba"
    try:
        ford = engine.lookup_dtc("P0101")
        expect(ford["status"] == "FORD SPECIFINIS KODAS", ford["status"])
        expect(ford["ford_note"] == "Testinė Ford pastaba", ford)
    finally:
        del engine.FORD_SPECIFIC_CODES["P0101"]

    suggestions = engine.suggest_codes("p01", limit=8)
    expect(suggestions and all(item["code"].startswith("P01") for item in suggestions), suggestions)
    expect(len(suggestions) <= 8, suggestions)

    print("engine checks passed")
    print("manufacturer sample", manufacturer)
    print("P0101", direct["description"])
    print("P0101 LT", direct["translation"])
    print("U0100", engine.lookup_dtc("U0100")["translation"])


if __name__ == "__main__":
    main()

"""DTC lookup, local technical translation, and indicative fault hints.

The dictionaries and translation rules are the notebook implementation.
"""
import re
from pathlib import Path
from docx import Document

DOCX_PATH = Path(__file__).resolve().parent / "DTC generatoriaus saltinis(2).docx"

# Čia įrašykite tik tuos kodus, kuriuos patvirtinote pagal Ford gamyklinę
# dokumentaciją konkrečiam modeliui / metams / varikliui.
# Pavyzdys: "P1234": "Trumpas Ford gamyklinis paaiškinimas"
FORD_SPECIFIC_CODES = {
    # "P1234": "Ford specifinis kodo paaiškinimas"
}

def extract_codes(docx_path):
    """Ištraukia DTC kodų ir jų aprašymų poras iš Word pastraipų bei lentelių."""
    if not docx_path.exists():
        raise FileNotFoundError(
            f"Nerastas Word failas: {docx_path.resolve()}\n"
            "Įkelkite .docx failą į tą patį aplanką arba pakeiskite DOCX_PATH."
        )

    doc = Document(str(docx_path))
    chunks = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            chunks.append(" | ".join(cell.text for cell in row.cells))

    full_text = "\n".join(chunks)

    # Atpažįsta formatus, pvz.: "P0101": "Mass Air Flow..."
    # Taip pat toleruoja kabutes ir kelių eilučių aprašymus.
    pattern = re.compile(
        r"""["']([PBCU]\d{4})["']\s*:\s*["'](.*?)["']""",
        flags=re.IGNORECASE | re.DOTALL
    )

    result = {}
    for match in pattern.finditer(full_text):
        code = match.group(1).upper()
        description = re.sub(r"\s+", " ", match.group(2)).strip()
        if description:
            result[code] = description
    return result

# Dalinis terminų žodynas. Originalus angliškas aprašymas visada rodomas,
# todėl netikslus automatinis vertimas neslepia šaltinio formuluotės.
TRANSLATIONS = [
    ("Crankshaft Position Sensor", "alkūninio veleno padėties jutiklis"),
    ("Camshaft Position Sensor", "skirstomojo veleno padėties jutiklis"),
    ("Engine Coolant Temperature", "variklio aušinimo skysčio temperatūra"),
    ("Intake Air Temperature", "įsiurbiamo oro temperatūra"),
    ("Fuel Rail/System Pressure", "degalų magistralės / sistemos slėgis"),
    ("Fuel Rail Pressure", "degalų magistralės slėgis"),
    ("Fuel Pressure Regulator", "degalų slėgio reguliatorius"),
    ("Mass or Volume Air flow", "oro masės / tūrio srautas"),
    ("Mass or Volume Air Flow", "oro masės / tūrio srautas"),
    ("Manifold Absolute Pressure", "įsiurbimo kolektoriaus absoliutusis slėgis"),
    ("Oxygen Sensor", "deguonies (lambda) jutiklis"),
    ("Heated Oxygen Sensor", "šildomas deguonies (lambda) jutiklis"),
    ("Glow Plug", "kaitinimo žvakė"),
    ("Injector Circuit", "purkštuko elektros grandinė"),
    ("Control Circuit", "valdymo grandinė"),
    ("Sensor Circuit", "jutiklio grandinė"),
    ("Heater Circuit", "šildytuvo grandinė"),
    ("Circuit", "grandinė"),
    ("Injector", "purkštukas"),
    ("Turbo / Super Charger", "turbokompresorius / kompresorius"),
    ("Turbocharger", "turbokompresorius"),
    ("Exhaust Gas Recirculation", "išmetamųjų dujų recirkuliacija (EGR)"),
    ("Fuel Pump", "degalų siurblys"),
    ("Fuel System", "degalų sistema"),
    ("Fuel Volume Regulator", "degalų kiekio reguliatorius"),
    ("Throttle Position", "droselinės sklendės padėtis"),
    ("Throttle", "droselinė sklendė"),
    ("Engine Oil Pressure", "variklio alyvos slėgis"),
    ("Engine Oil Temperature", "variklio alyvos temperatūra"),
    ("Vehicle Speed Sensor", "transporto priemonės greičio jutiklis"),
    ("Coolant Thermostat", "aušinimo skysčio termostatas"),
    ("Cooling Fan", "aušinimo ventiliatorius"),
    ("Crankshaft", "alkūninis velenas"),
    ("Camshaft", "skirstomasis velenas"),
    ("Circuit Malfunction", "elektros grandinės gedimas"),
    ("Range/Performance", "veikimo diapazono / našumo neatitiktis"),
    ("Performance", "veikimo / našumo sutrikimas"),
    ("Low Input", "per žemas įėjimo signalas"),
    ("High Input", "per aukštas įėjimo signalas"),
    ("Low Voltage", "per žema įtampa"),
    ("High Voltage", "per aukšta įtampa"),
    ("No Signal", "nėra signalo"),
    ("Intermittent/Erratic", "protarpinis / nestabilus signalas"),
    ("Intermittent", "protarpinis signalas"),
    ("Open", "nutraukta grandinė"),
    ("Shorted", "trumpasis jungimas"),
    ("Too Low", "per žemas"),
    ("Too High", "per aukštas"),
    ("Underboost", "per mažas pripūtimo slėgis"),
    ("Overboost", "per didelis pripūtimo slėgis"),
    ("Misfire Detected", "aptiktas degimo praleidimas"),
    ("Cylinder", "cilindras"),
    ("Bank 1", "1 cilindrų eilė"),
    ("Bank 2", "2 cilindrų eilė"),
    ("Sensor", "jutiklis"),
    ("System", "sistema"),
    ("Malfunction", "gedimas"),
    ("Detected", "aptiktas"),
]

# Išplėstas techninių frazių vertimo žodynas.
# Ilgesnės frazės keičiamos pirmiausia, kad žodžiai nebūtų verčiami atskirai
# ir nesusidarytų netaisyklingos konstrukcijos.
PHRASE_TRANSLATIONS = {
    # Tikslūs ratų greičio diagnostikos terminai – šios pilnos frazės
    # turi būti apdorojamos anksčiau už atskirus žodžius „speed“ ir „wheel“.
    "Speed Wheel Signal Mismatch": "Ratų greičio signalų neatitikimas",
    "Wheel Speed Signal Mismatch": "Ratų greičio signalų neatitikimas",
    "Speed Wheel Sensor Mismatch": "Ratų greičio jutiklių signalų neatitikimas",
    "Wheel Speed Sensor Mismatch": "Ratų greičio jutiklių signalų neatitikimas",
    "Front/Rear Wheel Speed Mismatch": "Priekinių ir galinių ratų greičio neatitikimas",
    "Front to Rear Wheel Speed Mismatch": "Priekinių ir galinių ratų greičio neatitikimas",
    "Left/Right Wheel Speed Mismatch": "Kairiojo ir dešiniojo ratų greičio neatitikimas",
    "Left to Right Wheel Speed Mismatch": "Kairiojo ir dešiniojo ratų greičio neatitikimas",
    "Speed Wheel Mismatch": "Ratų greičio neatitikimas",
    "Wheel Speed Mismatch": "Ratų greičio neatitikimas",
    "Wheel Speed Difference": "Ratų sukimosi greičio skirtumas",
    "Wheel Speed Comparison": "Ratų sukimosi greičio palyginimas",
    "Wheel Speed Sensor": "Rato sukimosi greičio jutiklis",
    "Speed Wheel Sensor": "Rato sukimosi greičio jutiklis",
    "Front Wheel Speed": "priekinių ratų sukimosi greitis",
    "Rear Wheel Speed": "galinių ratų sukimosi greitis",
    "Left Front Wheel Speed": "kairiojo priekinio rato sukimosi greitis",
    "Right Front Wheel Speed": "dešiniojo priekinio rato sukimosi greitis",
    "Left Rear Wheel Speed": "kairiojo galinio rato sukimosi greitis",
    "Right Rear Wheel Speed": "dešiniojo galinio rato sukimosi greitis",
    "Wheel Speed Sensor Circuit Low": "rato sukimosi greičio jutiklio grandinės signalas per žemas",
    "Wheel Speed Sensor Circuit High": "rato sukimosi greičio jutiklio grandinės signalas per aukštas",
    "Speed Wheel Sensor Circuit Low": "rato sukimosi greičio jutiklio grandinės signalas per žemas",
    "Speed Wheel Sensor Circuit High": "rato sukimosi greičio jutiklio grandinės signalas per aukštas",
    "Wheel Speed Sensor Circuit": "rato sukimosi greičio jutiklio elektros grandinė",
    "Wheel Speed Sensor Signal": "rato sukimosi greičio jutiklio signalas",
    "Wheel Speed Signal": "rato sukimosi greičio signalas",
    "Speed Sensor Signal": "greičio jutiklio signalas",
    "Wheel Speed": "ratų sukimosi greitis",
    "Mismatch": "neatitikimas",
    # Bendros diagnostinės formuluotės, kurios anksčiau likdavo angliškos.
    "Signal Plausibility Failure": "signalo patikimumo neatitiktis",
    "Plausibility Failure": "duomenų patikimumo neatitiktis",
    "Signal Out of Range": "signalo reikšmė nepatenka į leistiną diapazoną",
    "Data Out of Range": "duomenų reikšmė nepatenka į leistiną diapazoną",
    "Performance or Incorrect Operation": "veikimo sutrikimas arba netinkamas veikimas",
    "Incorrect Operation": "netinkamas veikimas",
    "Not Plausible": "neatitinka tikėtinos reikšmės",
    "Not Learned": "neatliktas adaptavimas",
    "Not Calibrated": "nesukalibruota",
    "Calibration Failed": "kalibravimas nepavyko",
    "Signal Compare Failure": "signalų palyginimo klaida",
    "Correlation Failure": "signalų tarpusavio neatitiktis",
    "Rationality Failure": "signalo loginės atitikties klaida",
    "Out of Range": "neleistino diapazono reikšmė",
    "Above Maximum Threshold": "viršyta didžiausia leistina riba",
    "Below Minimum Threshold": "nesiekiama mažiausia leistina riba",
    "Above Maximum": "viršija didžiausią leistiną reikšmę",
    "Below Minimum": "mažesnė už mažiausią leistiną reikšmę",
    "Open Load": "nutraukta apkrovos grandinė",
    "Short to Positive": "trumpasis jungimas į teigiamąjį maitinimo polių",
    "Short to Negative": "trumpasis jungimas į neigiamąjį maitinimo polių",
    "Short to Ground": "trumpasis jungimas į masę",
    "Open Circuit": "nutrūkusi elektros grandinė",
    "Circuit Open": "nutrūkusi elektros grandinė",
    "Intermittent Fault": "protarpinis gedimas",
    "Electrical Fault": "elektros gedimas",
    "Mechanical Fault": "mechaninis gedimas",
    "Internal Fault": "vidinis gedimas",
    "Supply Voltage": "maitinimo įtampa",
    "Reference Voltage": "etaloninė įtampa",
    "Reference Signal": "etaloninis signalas",
    "Ground Circuit": "masės grandinė",
    "Power Supply": "maitinimo šaltinis",
    "Ignition Voltage": "uždegimo sistemos maitinimo įtampa",
    "Battery Voltage": "akumuliatoriaus įtampa",
    "Engine Speed": "variklio sūkių dažnis",
    "Vehicle Speed": "transporto priemonės greitis",
    "Wheel Bearing": "rato guolis",
    "Steering Wheel": "vairas",
    "Steering Angle": "vairo pasukimo kampas",
    "Brake Pressure Sensor": "stabdžių slėgio jutiklis",
    "Brake Pedal Position": "stabdžių pedalo padėtis",
    "Accelerator Pedal Position": "akceleratoriaus pedalo padėtis",
    "Throttle Position Sensor": "droselinės sklendės padėties jutiklis",
    "Intake Manifold": "įsiurbimo kolektorius",
    "Exhaust Manifold": "išmetimo kolektorius",
    "Exhaust Temperature": "išmetamųjų dujų temperatūra",
    "Particulate Filter": "kietųjų dalelių filtras",
    "Diesel Particulate Filter": "dyzelino kietųjų dalelių filtras (DPF)",
    "Differential Pressure": "slėgių skirtumas",
    "Pressure Sensor": "slėgio jutiklis",
    "Temperature Sensor": "temperatūros jutiklis",
    "Position Sensor": "padėties jutiklis",
    "Circuit Low": "per žemas elektros grandinės signalas",
    "Circuit High": "per aukštas elektros grandinės signalas",
    "Circuit Range/Performance": "elektros grandinės veikimo diapazono / našumo neatitiktis",
    "No Signal": "signalo nėra",
    "Signal Missing": "trūksta signalo",
    "Signal Invalid": "netinkamas signalas",
    "Invalid Signal": "netinkamas signalas",
    "Data Invalid": "netinkami duomenys",
    "Invalid Data": "netinkami duomenys",
    "Lost Communication With": "Prarastas ryšys su",
    "No Communication With": "Nėra ryšio su",
    # U serija – tinklo ir valdymo blokų ryšys
    "Lost Communication With": "Prarastas ryšys su",
    "Lost Communication": "Prarastas ryšys",
    "No Communication With": "Nėra ryšio su",
    "Communication With": "Ryšys su",
    "High Speed CAN Communication Bus": "Didelės spartos CAN ryšio magistralė",
    "Medium Speed CAN Communication Bus": "Vidutinės spartos CAN ryšio magistralė",
    "Low Speed CAN Communication Bus": "Mažos spartos CAN ryšio magistralė",
    "CAN Communication Bus": "CAN ryšio magistralė",
    "CAN Bus": "CAN magistralė",
    "Control Module": "valdymo blokas",
    "Engine Control Module": "variklio valdymo blokas",
    "Transmission Control Module": "pavarų dėžės valdymo blokas",
    "Anti-Lock Brake System": "stabdžių antiblokavimo sistema (ABS)",
    "Brake System Control Module": "stabdžių sistemos valdymo blokas",
    "Power Steering Control Module": "vairo stiprintuvo valdymo blokas",
    "Steering Angle Sensor Module": "vairo kampo jutiklio blokas",
    "Tire Pressure Monitor Module": "padangų slėgio stebėjimo sistemos valdymo blokas",
    "Vehicle Dynamics Control Module": "automobilio dinamikos valdymo blokas",
    "Yaw Rate Sensor Module": "sukimosi apie vertikaliąją ašį greičio jutiklio blokas",
    "Lateral Acceleration Sensor Module": "šoninio pagreičio jutiklio blokas",
    "Multi-axis Acceleration Sensor Module": "daugi ašių pagreičio jutiklio blokas",
    "Gear Shift Module": "pavarų perjungimo valdymo blokas",
    "Cruise Control Module": "pastovaus greičio palaikymo sistemos valdymo blokas",
    "Fuel Injector Control Module": "degalų purkštukų valdymo blokas",
    "Glow Plug Control Module": "kaitinimo žvakių valdymo blokas",
    "Throttle Actuator Control Module": "droselinės sklendės pavaros valdymo blokas",
    "Fuel Pump Control Module": "degalų siurblio valdymo blokas",
    "Drive Motor Control Module": "traukos variklio valdymo blokas",
    "Battery Energy Control Module": "akumuliatoriaus energijos valdymo blokas",
    "Transfer Case Control Module": "skirstomosios pavarų dėžės valdymo blokas",
    "Four-Wheel Drive Clutch Control Module": "visų varančiųjų ratų sankabos valdymo blokas",
    "Coolant Temperature Control Module": "aušinimo skysčio temperatūros valdymo blokas",
    "Starter / Generator Control Module": "starterio ir generatoriaus valdymo blokas",
    "Park Brake Control Module": "stovėjimo stabdžio valdymo blokas",
    "Steering Effort Control Module": "vairo pasipriešinimo valdymo blokas",
    "Ride Level Control Module": "kėbulo aukščio reguliavimo valdymo blokas",
    "Active Roll Control Module": "aktyviosios kėbulo posvyrio kontrolės valdymo blokas",
    "Electronic Brake Control Module": "elektroninis stabdžių valdymo blokas",
    "Body Control Module": "kėbulo elektronikos valdymo blokas",
    "Instrument Panel Control Module": "prietaisų skydelio valdymo blokas",
    "Airbag Control Module": "oro pagalvių valdymo blokas",
    "Restraints Control Module": "pasyviosios saugos sistemos valdymo blokas",
    "Door Control Module": "durelių valdymo blokas",
    "Seat Control Module": "sėdynės valdymo blokas",
    "Parking Assist Control Module": "parkavimo sistemos valdymo blokas",
    "Radio Control Module": "radijo imtuvo valdymo blokas",
    "Audio Control Module": "garso sistemos valdymo blokas",
    "Navigation Control Module": "navigacijos sistemos valdymo blokas",
    "HVAC Control Module": "klimato kontrolės valdymo blokas",
    "ECM/PCM A": "variklio / jėgos pavaros valdymo blokas A",
    "ECM/PCM B": "variklio / jėgos pavaros valdymo blokas B",
    "ECM/PCM": "variklio / jėgos pavaros valdymo blokas",
    "PCM": "jėgos pavaros valdymo blokas (PCM)",
    "ECM": "variklio valdymo blokas (ECM)",
    "TCM": "pavarų dėžės valdymo blokas (TCM)",
    "ABS": "stabdžių antiblokavimo sistema (ABS)",
    "Invalid Data Received From": "Gauti neteisingi duomenys iš",
    "Missing Message": "Trūksta pranešimo iš",
    "Message Missing": "Trūksta pranešimo",
    "Data Bus": "duomenų magistralė",
    "Bus Off": "magistralė neveikia (Bus Off)",
    "Bus Communication": "magistralės ryšys",
    "Network Communication": "tinklo ryšys",
    "Control Module Communication": "valdymo blokų ryšys",
    "Communication Circuit": "ryšio grandinė",
    "Communication Error": "ryšio klaida",
    "Network": "duomenų tinklas",

    # B serija – kėbulas, komforto ir pasyviosios saugos sistemos
    "Right Front Passenger Side": "dešinės priekinės keleivio pusės",
    "Left Front Driver Side": "kairės priekinės vairuotojo pusės",
    "Deployment Loop Resistance": "oro pagalvės išskleidimo grandinės varža",
    "Climate Control": "klimato kontrolės sistema",
    "Pushbutton": "mygtukas",
    "Fuel Sender": "degalų lygio jutiklis",
    "Fuel Tank Pressure Sensor": "degalų bako slėgio jutiklis",
    "Fuel Temperature Sensor": "degalų temperatūros jutiklis",
    "Anti-Theft": "apsaugos nuo vagystės sistema",
    "Programmed Keys": "užprogramuoti raktai",
    "Running Board Lamp": "šoninio laiptelio apšvietimo žibintas",
    "Emergency & Road Side Assistance Switch": "avarinės pagalbos iškvietimo jungiklis",
    "Horn Relay Coil": "garso signalo relės ritė",
    "Glass Break Sensor": "stiklo dūžio jutiklis",
    "Mirror Switch": "veidrodėlių valdymo jungiklis",
    "Window Feedback": "elektra valdomo lango grįžtamojo ryšio signalas",
    "Seat Belt": "saugos diržas",
    "Air Bag": "oro pagalvė",
    "Airbag": "oro pagalvė",
    "Occupant Classification": "keleivio atpažinimo sistema",
    "Passenger Presence": "keleivio buvimo jutiklis",
    "Door Ajar": "atidarytų durelių signalas",
    "Central Locking": "centrinis užraktas",
    "Power Window": "elektra valdomas langas",
    "Sunroof": "stoglangis",
    "Wiper Motor": "stiklo valytuvų variklis",
    "Washer Pump": "stiklų plovimo siurblys",
    "Turn Signal": "posūkio rodiklis",
    "Brake Lamp": "stabdžių žibintas",
    "Tail Lamp": "galinis žibintas",
    "Headlamp": "priekinis žibintas",
    "Interior Lamp": "salono apšvietimo žibintas",
    "Seat Position": "sėdynės padėtis",
    "Restraint System": "pasyviosios saugos sistema",
    "Deployment Loop": "oro pagalvės išskleidimo grandinė",
    "Resistance Low": "per maža varža",
    "Resistance High": "per didelė varža",
    "Number of Programmed Keys Is Below Minimum": "užprogramuotų raktų skaičius mažesnis už minimalų leidžiamą skaičių",
    "Longitudinal Acceleration Threshold Exceeded": "viršyta išilginio pagreičio riba",
    "See Manufacturer": "žr. gamintojo techninę dokumentaciją",
    "Invalid Code": "neteisingas kodas",

    # C serija – važiuoklė, stabdžiai ir judėjimo sistemos
    "Speed Wheel Sensor": "rato greičio jutiklis",
    "Wheel Speed Sensor": "rato sukimosi greičio jutiklis",
    "Tone Ring": "impulsinis žiedas",
    "Missing Tooth": "trūksta danties",
    "Coherency Failure": "signalų tarpusavio neatitiktis",
    "Hydraulic Pump Motor": "hidraulinio siurblio elektros variklis",
    "ABS Hydraulic Pump": "ABS hidraulinis siurblys",
    "ABS Pump": "ABS siurblys",
    "Hydraulic Valve": "hidraulinis vožtuvas",
    "Acceleration Switch": "pagreičio jungiklis",
    "Hydraulic Brake Switch": "hidraulinio stabdžių slėgio jungiklis",
    "Traction Control": "traukos kontrolės sistema",
    "Disable Lamp": "išjungimo indikatoriaus lemputė",
    "Disable Switch": "išjungimo jungiklis",
    "Function Enabled Input": "funkcijos įjungimo įėjimo signalas",
    "Speed Control Actuator Assembly": "pastovaus greičio palaikymo sistemos vykdiklio mazgas",
    "Cable Release": "troselio atleidimas",
    "Power Relay Coil": "maitinimo relės ritė",
    "Power Relay Output": "maitinimo relės išėjimas",
    "Starter Motor": "starterio elektros variklis",
    "RPM Input": "variklio sūkių dažnio įėjimo signalas",
    "Input Shaft Speed": "įėjimo veleno sukimosi greitis",
    "Brake Fluid Level Sensor": "stabdžių skysčio lygio jutiklis",
    "Cruise Control Command Switch": "pastovaus greičio palaikymo sistemos valdymo jungiklis",
    "Cruise Control Deactivator Brake Switch": "pastovaus greičio palaikymo sistemos išjungimo stabdžių jungiklis",
    "Clutch Position": "sankabos padėtis",
    "Gear Shift Position": "pavarų perjungimo padėtis",
    "Gear Select Position": "pasirinktos pavaros padėtis",
    "Steering Angle": "vairo pasukimo kampas",
    "Yaw Rate": "sukimosi apie vertikaliąją ašį greitis",
    "Lateral Acceleration": "šoninis pagreitis",
    "Brake Pressure": "stabdžių slėgis",
    "ABS": "stabdžių antiblokavimo sistema (ABS)",
    "Traction": "traukos",
    "Vehicle Speed": "transporto priemonės greitis",

    # P serija – variklis ir jėgos pavara
    "Fuel Volume Regulator": "degalų kiekio reguliatorius",
    "Fuel Shutoff Valve": "degalų uždarymo vožtuvas",
    "Fuel Rail Pressure": "degalų magistralės slėgis",
    "Fuel Rail": "degalų magistralė",
    "Fuel Pressure": "degalų slėgis",
    "Fuel Pump": "degalų siurblys",
    "Fuel System": "degalų sistema",
    "Intake Camshaft": "įsiurbimo skirstomasis velenas",
    "Exhaust Camshaft": "išmetimo skirstomasis velenas",
    "Camshaft Position Actuator": "skirstomojo veleno padėties reguliatorius",
    "Camshaft Position": "skirstomojo veleno padėtis",
    "Crankshaft Position": "alkūninio veleno padėtis",
    "Engine Position System": "variklio velenų padėties sistema",
    "Oxygen Sensor": "deguonies (lambda) jutiklis",
    "Heated Oxygen Sensor": "šildomas deguonies (lambda) jutiklis",
    "Catalyst System": "katalizatoriaus sistema",
    "Catalyst": "katalizatorius",
    "Exhaust Gas Recirculation": "išmetamųjų dujų recirkuliacija (EGR)",
    "EGR": "išmetamųjų dujų recirkuliacijos sistema (EGR)",
    "Mass or Volume Air Flow": "oro masės arba tūrio srautas",
    "Mass Air Flow": "oro masės srautas",
    "Manifold Absolute Pressure": "įsiurbimo kolektoriaus absoliutusis slėgis",
    "Intake Air Temperature": "įsiurbiamo oro temperatūra",
    "Engine Coolant Temperature": "variklio aušinimo skysčio temperatūra",
    "Coolant Temperature": "aušinimo skysčio temperatūra",
    "Throttle Position": "droselinės sklendės padėtis",
    "Throttle Actuator": "droselinės sklendės pavara",
    "Glow Plug": "kaitinimo žvakė",
    "Injector Circuit": "purkštuko elektros grandinė",
    "Fuel Injector": "degalų purkštukas",
    "Turbo/Super Charger": "turbokompresorius arba mechaninis kompresorius",
    "Turbocharger": "turbokompresorius",
    "Boost Pressure": "pripūtimo slėgis",
    "Engine Oil Pressure": "variklio alyvos slėgis",
    "Engine Oil Temperature": "variklio alyvos temperatūra",
    "Vehicle Speed Sensor": "transporto priemonės greičio jutiklis",
    "Misfire Detected": "aptiktas degimo praleidimas",
    "Random/Multiple Cylinder": "atsitiktinis / kelių cilindrų",
    "Cylinder": "cilindras",
    "Bank 1": "1-oji cilindrų eilė",
    "Bank 2": "2-oji cilindrų eilė",

    # Bendros diagnostinės formuluotės
    "Control Circuit": "valdymo elektros grandinė",
    "Sensor Circuit": "jutiklio elektros grandinė",
    "Heater Control Circuit": "šildytuvo valdymo elektros grandinė",
    "Heater Circuit": "šildytuvo elektros grandinė",
    "Circuit / Open": "elektros grandinė nutrūkusi",
    "Open Circuit": "nutrūkusi elektros grandinė",
    "Circuit Open": "elektros grandinė nutrūkusi",
    "Short To Ground": "trumpasis jungimas į masę",
    "Short to Ground": "trumpasis jungimas į masę",
    "Short To Battery": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Short to Battery": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Short to Vbatt": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Short Circuit To Ground": "trumpasis jungimas į masę",
    "Short Circuit To Battery": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Circuit Failure": "elektros grandinės gedimas",
    "Circuit Malfunction": "elektros grandinės veikimo sutrikimas",
    "Circuit Range/Performance": "elektros grandinės veikimo diapazono / našumo neatitiktis",
    "Range/Performance": "veikimo diapazono / našumo neatitiktis",
    "Low Voltage": "per žema įtampa",
    "High Voltage": "per aukšta įtampa",
    "Low Input": "per žemas įėjimo signalas",
    "High Input": "per aukštas įėjimo signalas",
    "Signal Low": "per žemas signalo lygis",
    "Signal High": "per aukštas signalo lygis",
    "No Signal": "signalo nėra",
    "Signal Missing": "trūksta signalo",
    "Intermittent/Erratic": "protarpinis / nepastovus signalas",
    "Intermittent": "protarpinis",
    "Performance": "veikimo / našumo sutrikimas",
    "Over-Advanced": "pernelyg paankstintas",
    "Over-Retarded": "pernelyg suvėlintas",
    "Too Lean": "per liesas mišinys",
    "Too Rich": "per riebus mišinys",
    "Underboost": "per mažas pripūtimo slėgis",
    "Overboost": "per didelis pripūtimo slėgis",
    "Insufficient": "nepakankamas",
    "Excessive": "per didelis",
    "Low": "per žemas",
    "High": "per aukštas",
    "Open": "nutrūkusi grandinė",
    "Short": "trumpasis jungimas",
    "Range": "diapazonas",
    "Input": "įėjimo signalas",
    "Output": "išėjimo signalas",
    "Voltage": "įtampa",
    "Sensor": "jutiklis",
    "Switch": "jungiklis",
    "Relay": "relė",
    "Motor": "elektros variklis",
    "Pump": "siurblys",
    "Valve": "vožtuvas",
    "Actuator": "vykdiklis",
    "System": "sistema",
    "Failure": "gedimas",
    "Malfunction": "veikimo sutrikimas",
    "Detected": "aptiktas",
    "Not Detected": "neaptiktas",
    "Missing": "trūksta",
    "Invalid": "neteisingas",
    "Position": "padėtis",
    "Temperature": "temperatūra",
    "Pressure": "slėgis",
    "Speed": "greitis",
    "Level": "lygis",
    "Circuit": "elektros grandinė",
    "Engine": "variklis",
    "Transmission": "pavarų dėžė",
    "Vehicle": "transporto priemonė",
}


# Papildytas terminų žodynas pagal realius 3 445 skirtingus šaltinio aprašymus.
# Terminai parinkti pagal diagnostinę reikšmę; santrumpos ir gamintojo žymos išsaugomos.
EXTRA_PHRASE_TRANSLATIONS = {
    "Evaporative Emission Control System": "degalų garų kontrolės sistema (EVAP)",
    "Evaporative Emission Control": "degalų garų kontrolė (EVAP)",
    "Evaporative Emission": "degalų garų emisija",
    "Air Bag Driver": "vairuotojo oro pagalvė",
    "Air Bag Passenger": "keleivio oro pagalvė",
    "Driver Air Bag": "vairuotojo oro pagalvė",
    "Passenger Air Bag": "keleivio oro pagalvė",
    "Driver Side": "vairuotojo pusės",
    "Passenger Side": "keleivio pusės",
    "Front Side": "priekinės pusės",
    "Rear Side": "galinės pusės",
    "Accelerometer Front": "priekinis pagreičio jutiklis",
    "Accelerometer Rear": "galinis pagreičio jutiklis",
    "ABS Inlet Valve": "ABS įleidimo vožtuvas",
    "ABS Outlet Valve": "ABS išleidimo vožtuvas",
    "Inlet Valve Coil": "įleidimo vožtuvo ritė",
    "Outlet Valve Coil": "išleidimo vožtuvo ritė",
    "Blower (Fan)": "salono ventiliatorius",
    "Blower Fan": "salono ventiliatorius",
    "Air Suspension": "pneumatinė pakaba",
    "Air Suspension Compressor": "pneumatinės pakabos kompresorius",
    "Power Sliding Door": "elektra valdoma slankioji durelė",
    "Sliding Door": "slankioji durelė",
    "B-pillar": "B statramstis",
    "A/C Clutch": "oro kondicionieriaus kompresoriaus sankaba",
    "A/C Evaporator": "oro kondicionieriaus garintuvas",
    "A/C Post Heater": "oro kondicionieriaus papildomas šildytuvas",
    "Air Conditioner": "oro kondicionierius",
    "Air Conditioning": "oro kondicionavimo sistema",
    "Refrigerant Charge Loss": "šaltnešio kiekio sumažėjimas",
    "Refrigerant": "šaltnešis",
    "Anti Theft": "apsaugos nuo vagystės sistema",
    "Anti-Theft": "apsaugos nuo vagystės sistema",
    "Theft Deterrent": "apsaugos nuo vagystės sistema",
    "Horn Output": "garso signalo išėjimas",
    "Output Circuit": "išėjimo elektros grandinė",
    "Input Circuit": "įėjimo elektros grandinė",
    "Input Signal": "įėjimo signalas",
    "Output Signal": "išėjimo signalas",
    "Signal Circuit": "signalo elektros grandinė",
    "Signal Is Not Sensed": "signalas neaptinkamas",
    "Is Not Sensed": "neaptinkamas",
    "Out of Self Test Range": "savitikros diapazono neatitiktis",
    "Self Test Range": "savitikros diapazonas",
    "Out of Limits": "už leistinų ribų",
    "Out of Range": "už leistino diapazono ribų",
    "Range/Performance": "veikimo diapazono / charakteristikų neatitiktis",
    "Range / Performance": "veikimo diapazono / charakteristikų neatitiktis",
    "Short To Batt": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Short to Batt": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Short To Gnd": "trumpasis jungimas į masę",
    "Short to Gnd": "trumpasis jungimas į masę",
    "Short To Ground": "trumpasis jungimas į masę",
    "Short to Ground": "trumpasis jungimas į masę",
    "Short To Battery": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Short to Battery": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Short To Vbatt": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Short to Vbatt": "trumpasis jungimas į akumuliatoriaus teigiamąjį polių",
    "Ground Short": "trumpasis jungimas į masę",
    "Battery Short": "trumpasis jungimas į akumuliatoriaus maitinimo grandinę",
    "Circuit Failure": "elektros grandinės gedimas",
    "Circuit Fault": "elektros grandinės gedimas",
    "Circuit Malfunction": "elektros grandinės veikimo sutrikimas",
    "Circuit Open": "elektros grandinė nutrūkusi",
    "Open Circuit": "nutrūkusi elektros grandinė",
    "Short Circuit": "trumpasis jungimas",
    "Resistance Low": "per maža varža",
    "Resistance High": "per didelė varža",
    "Potentiometer Circuit": "potenciometro elektros grandinė",
    "Potentiometer": "potenciometras",
    "Solenoid Control": "solenoido valdymas",
    "Solenoid": "solenoidas",
    "Coil Circuit": "ritės elektros grandinė",
    "Coil": "ritė",
    "Shift Solenoid": "pavarų perjungimo solenoidas",
    "Shift": "perjungimas",
    "Gear Shift": "pavarų perjungimas",
    "Transmission": "pavarų dėžė",
    "Transfer Case": "skirstomoji dėžė",
    "Power Steering": "vairo stiprintuvas",
    "Steering Wheel": "vairas",
    "Steering Angle": "vairo pasukimo kampas",
    "Ride Height": "pakabos aukštis",
    "Height Sensor": "aukščio jutiklis",
    "Seat Belt": "saugos diržas",
    "Seatbelt": "saugos diržas",
    "Seat Position": "sėdynės padėtis",
    "Seat Recline": "sėdynės atlošo padėtis",
    "Child Seat": "vaikiška automobilinė kėdutė",
    "Occupant Classification": "keleivio klasifikavimo sistema",
    "Passenger Presence": "keleivio buvimo sistema",
    "Air Bag": "oro pagalvė",
    "Airbag": "oro pagalvė",
    "Inflator": "oro pagalvės pripūtimo įtaisas",
    "Squib": "oro pagalvės uždegiklis",
    "Restraint System": "pasyviosios saugos sistema",
    "Restraints": "pasyviosios saugos sistema",
    "Wiper Motor": "stiklo valytuvų elektros variklis",
    "Washer Pump": "stiklų plovimo siurblys",
    "Turn Signal": "posūkio rodiklis",
    "Brake Lamp": "stabdžių žibintas",
    "Tail Lamp": "galinis žibintas",
    "Headlamp": "priekinis žibintas",
    "Courtesy Lamp": "salono apšvietimo lemputė",
    "Dome Lamp": "salono plafono lemputė",
    "Warning Lamp": "įspėjamoji lemputė",
    "Indicator Lamp": "indikatoriaus lemputė",
    "Mirror Switch": "veidrodėlių valdymo jungiklis",
    "Power Window": "elektra valdomas langas",
    "Window Motor": "lango pakėlimo elektros variklis",
    "Door Ajar": "atidarytų durelių būsena",
    "Door Lock": "durelių užraktas",
    "Central Locking": "centrinio užrakto sistema",
    "Cruise Control": "pastovaus greičio palaikymo sistema",
    "Brake Pedal": "stabdžių pedalas",
    "Accelerator Pedal": "akceleratoriaus pedalas",
    "Clutch Pedal": "sankabos pedalas",
    "Brake Fluid": "stabdžių skystis",
    "Hydraulic Pump": "hidraulinis siurblys",
    "Hydraulic Valve": "hidraulinis vožtuvas",
    "Fuel Rail": "degalų magistralė",
    "Fuel Pressure": "degalų slėgis",
    "Fuel Pump": "degalų siurblys",
    "Fuel Injector": "degalų purkštukas",
    "Fuel Tank": "degalų bakas",
    "Fuel Level": "degalų lygis",
    "Fuel Volume": "degalų kiekis",
    "Fuel Control": "degalų mišinio valdymas",
    "Fuel Trim": "degalų mišinio korekcija",
    "Air Flow": "oro srautas",
    "Air Intake": "oro įsiurbimas",
    "Intake Air": "įsiurbiamas oras",
    "Exhaust Gas": "išmetamosios dujos",
    "Exhaust Temperature": "išmetamųjų dujų temperatūra",
    "Coolant Temperature": "aušinimo skysčio temperatūra",
    "Engine Coolant": "variklio aušinimo skystis",
    "Engine Oil": "variklio alyva",
    "Oil Pressure": "alyvos slėgis",
    "Oil Temperature": "alyvos temperatūra",
    "Glow Plug": "kaitinimo žvakė",
    "Aux Heater": "papildomas šildytuvas",
    "Heater": "šildytuvas",
    "Cooling Fan": "aušinimo ventiliatorius",
    "Additional Fan": "papildomas ventiliatorius",
    "Fan Relay": "ventiliatoriaus relė",
    "Blower": "salono ventiliatorius",
    "Evaporator": "garintuvas",
    "Blend Door": "oro srauto maišymo sklendė",
    "Recirculation Door": "oro recirkuliacijos sklendė",
    "Air Bypass": "oro apėjimo kanalas",
    "Bypass Valve": "apėjimo vožtuvas",
    "Throttle": "droselinė sklendė",
    "Camshaft": "skirstomasis velenas",
    "Crankshaft": "alkūninis velenas",
    "Cam Control": "skirstomojo veleno valdymas",
    "Timing": "fazės / sinchronizavimo laikas",
    "Misfire": "degimo praleidimas",
    "Cylinder": "cilindras",
    "Bank": "cilindrų eilė",
    "Front": "priekinis",
    "Rear": "galinis",
    "Left": "kairysis",
    "Right": "dešinysis",
    "Driver": "vairuotojo",
    "Passenger": "keleivio",
    "Battery": "akumuliatorius",
    "Power": "maitinimas",
    "Ground": "masė",
    "Voltage": "įtampa",
    "Current": "srovė",
    "Temperature": "temperatūra",
    "Pressure": "slėgis",
    "Speed": "greitis",
    "Position": "padėtis",
    "Level": "lygis",
    "Switch": "jungiklis",
    "Relay": "relė",
    "Motor": "elektros variklis",
    "Pump": "siurblys",
    "Valve": "vožtuvas",
    "Actuator": "vykdiklis",
    "Module": "valdymo modulis",
    "Communication": "ryšys",
    "Data": "duomenys",
    "Signal": "signalas",
    "Fault": "gedimas",
    "Failure": "gedimas",
    "Malfunction": "veikimo sutrikimas",
    "Missing": "trūksta",
    "Invalid": "neteisingas",
    "Detected": "aptiktas",
    "Not Detected": "neaptiktas",
    "Intermittent": "protarpinis",
    "Low": "per žemas",
    "High": "per aukštas",
    "Open": "nutrūkusi grandinė",
    "Short": "trumpasis jungimas",
    "Output": "išėjimo signalas",
    "Input": "įėjimo signalas",
    "Rearward": "galinis / atgalinis",
    "Forward": "priekinis / pirmyn",
    "Up/Down": "aukštyn / žemyn",
    "Up": "aukštyn",
    "Down": "žemyn",
    "Stuck": "užstrigęs",
    "Leak": "nuotėkis",
    "Excessive": "per didelis",
    "Insufficient": "nepakankamas",
    "Failed": "nepavyko",
    "Error": "klaida",
    "Request": "užklausa",
    "Feedback": "grįžtamasis ryšys",
    "Reference": "etaloninis",
    "Supply": "maitinimas",
    "Ignition": "uždegimas",
    "Starter": "starteris",
    "Alternator": "generatorius",
    "Charging System": "įkrovimo sistema",
    "Emission": "išmetamųjų teršalų kontrolė",
    "Catalyst": "katalizatorius",
    "Particulate Filter": "kietųjų dalelių filtras",
    "Diesel Particulate Filter": "dyzelino kietųjų dalelių filtras (DPF)",
    "EGR": "išmetamųjų dujų recirkuliacija (EGR)",
    "Oxygen Sensor": "deguonies (lambda) jutiklis",
    "O2 Sensor": "deguonies (lambda) jutiklis",
    "HO2S": "šildomas deguonies (lambda) jutiklis",
    "MAF": "oro masės srauto jutiklis (MAF)",
    "MAP": "įsiurbimo kolektoriaus slėgio jutiklis (MAP)",
    "Gnd": "masė",
    "Batt": "akumuliatoriaus teigiamasis polius",
    "Vbatt": "akumuliatoriaus įtampa",
    "Ckt": "elektros grandinė",
    "Temp": "temperatūra",
    "A/C": "oro kondicionavimo sistema",
    "Battery Temperature": "akumuliatoriaus temperatūra",
    "Ambient Air Temperature": "aplinkos oro temperatūra",
    "Air Temperature External": "išorės oro temperatūra",
    "Air Temperature Internal": "salono oro temperatūra",
    "Air Temperature": "oro temperatūra",
    "Climate Control": "klimato kontrolės sistema",
    "Air Assist": "papildomas oro tiekimas",
    "Air Assisted": "oru papildomai valdomas",
    "Air Mixture": "oro mišinio",
    "Air Pump": "antrinio oro siurblys",
    "Air Spring": "pneumatinė spyruoklė",
    "Air Compress": "oro kompresorius",
    "Air Flow": "oro srautas",
    "Air": "oras",
    "Control": "valdymo",
    "Clutch": "sankabos",
    "Door": "durelių",
    "Lamp": "lemputės",
    "Brake": "stabdžių",
    "Seat": "sėdynės",
    "Side": "pusės",
    "Fuel": "degalų",
    "Battery": "akumuliatoriaus",
    "Front": "priekinės dalies",
    "Rear": "galinės dalies",
    "Left": "kairės pusės",
    "Right": "dešinės pusės",
    "Driver": "vairuotojo",
    "Passenger": "keleivio",
    "External": "išorinis",
    "Internal": "vidinis",
    "Ambient": "aplinkos",
    "Aux": "papildomas",
    "To": "į",
    "From": "iš",
    "With": "su",
    "Or": "arba",
    "And": "ir",
    "Not": "ne",
    "Responding": "atsako",
    "Exceeds": "viršija",
    "Exceeded": "viršyta",
    "Maximum": "didžiausia leistina",
    "Minimum": "mažiausia leistina",
    "Max": "maksimali",
    "Min": "minimali",
    "Dumping": "slėgio išleidimas",
    "Fluid": "skystis",
    "Upshift": "aukštesnės pavaros įjungimas",
    "Skip Shift": "pavaros praleidimas",
    "High To Low Side": "aukštojo ir žemojo potencialo pusės",
    "High To Low Side Open": "aukštojo ir žemojo potencialo pusės grandinės nutrūkimas",
    "High To Low Side Short": "aukštojo ir žemojo potencialo pusių trumpasis jungimas",
    "Power Supply": "maitinimo šaltinis",
    "B+ Supply": "teigiamojo maitinimo šaltinis",
    "VCRM": "kintamo greičio valdymo modulis (VCRM)",
    "KOEO": "uždegimas įjungtas, variklis neveikia (KOEO)",
    "KOER": "variklis veikia (KOER)",
    "KAM": "nuolatinė adaptyvioji atmintis (KAM)",
    "RAM/ROM": "RAM/ROM atmintis",
    "TRS": "pavarų dėžės diapazono jutiklis (TRS)",
    "LF": "kairysis priekinis",
    "RF": "dešinysis priekinis",
    "LR": "kairysis galinis",
    "RR": "dešinysis galinis",
    "Exhaust Camshaft": "išmetimo skirstomasis velenas",
    "Intake Camshaft": "įsiurbimo skirstomasis velenas",
    "Exhaust Valve": "išmetimo vožtuvas",
    "Intake Valve": "įsiurbimo vožtuvas",
    "Exhaust Pressure": "išmetamųjų dujų slėgis",
    "Exhaust Not Warm": "išmetimo sistema neįšilusi",
    "Downstream O2 Sensor": "už katalizatoriaus esantis deguonies (lambda) jutiklis",
    "Upstream O2 Sensor": "prieš katalizatorių esantis deguonies (lambda) jutiklis",
    "External Charging Fault": "išorinis įkrovimo sistemos gedimas",
    "Charging Fault": "įkrovimo sistemos gedimas",
    "Charging": "įkrovimas",
    "Fan Control": "ventiliatoriaus valdymas",
    "Fan Speed": "ventiliatoriaus sukimosi greitis",
    "Fan Secondary": "antrinio ventiliatoriaus",
    "Fan Overspeed": "per didelis ventiliatoriaus sukimosi greitis",
    "High Fan": "didelis ventiliatoriaus greitis",
    "Low Fan": "mažas ventiliatoriaus greitis",
    "Electrodrive Fan": "elektrinis ventiliatorius",
    "Fan": "ventiliatorius",
    "Interior Fan": "salono ventiliatorius",
    "Gear Ratio": "pavaros perdavimo santykis",
    "Incorrect Gear Ratio": "netinkamas pavaros perdavimo santykis",
    "Incorrect Ratio": "netinkamas perdavimo santykis",
    "Gear Lever": "pavarų svirtis",
    "Gear Control": "pavarų valdymas",
    "First Gear": "pirmoji pavara",
    "Gear": "pavara",
    "Loss of KAM Power": "nuolatinės adaptyviosios atminties (KAM) maitinimo praradimas",
    "Loss of Pressure": "slėgio praradimas",
    "Lack Of": "trūksta",
    "Lack of": "trūksta",
    "Not Responding": "neatsako",
    "Not Warm": "neįšilęs",
    "Sensor Indicates Lean": "jutiklis rodo liesą mišinį",
    "Sensor Indicates Rich": "jutiklis rodo riebų mišinį",
    "Indicates Lean": "rodo liesą mišinį",
    "Indicates Rich": "rodo riebų mišinį",
    "Adaptive Fuel": "degalų mišinio adaptacija",
    "At Limit": "pasiekė ribinę reikšmę",
    "Overspeed": "per didelis sukimosi greitis",
    "Clutch Locked": "sankaba užblokuota",
    "Clutch": "sankaba",
    "Locked": "užblokuotas",
    "Secondary": "antrinis",
    "Primary": "pagrindinis",
    "Off": "išjungta",
    "On": "įjungta",
    "Warm": "įšilęs",
    "Light": "lemputė",
    "Monitor": "stebėsena",
    "Operational": "veikimo",
    "Spring/Shock": "pneumatinė spyruoklė / amortizatorius",
    "Air Spring/Shock": "pneumatinė spyruoklė / amortizatorius",
    "Air Spring": "pneumatinė spyruoklė",
    "Shock": "amortizatorius",
    "Spring": "spyruoklė",
    "Pressure Reduction": "slėgio sumažinimas",
    "Reduction": "sumažinimas",
    "Problem": "problema",
    "Hydraulic Pressure Unit": "hidraulinis slėgio blokas",
    "Max Timing": "maksimalus laikas",
    "Request Exceeded Max Timing": "užklausos laikas viršijo maksimalią ribą",
    "VCRM Fan": "VCRM ventiliatorius",
    "Engine Temperature": "variklio temperatūra",
    "Temperature Light": "temperatūros įspėjamoji lemputė",
    "First Gear Switch": "pirmosios pavaros jungiklis",
    "Four Wheel Drive": "keturių varančiųjų ratų pavara (4WD)",
    "Left Front Wheel": "kairysis priekinis ratas",
    "Left Rear Wheel": "kairysis galinis ratas",
    "Right Front Wheel": "dešinysis priekinis ratas",
    "Right Rear Wheel": "dešinysis galinis ratas",
    "Wheel Pressure": "padangos slėgis",
    "Wheel": "ratas",
    "Drive": "pavara",
    "Power": "maitinimas",
    "Air Spring Solenoid": "pneumatinės spyruoklės solenoidas",
    "Solenoid Output": "solenoido išėjimo grandinė",
    "Operational Failure": "veikimo sutrikimas",
    "Transition States": "būsenų perjungimas",
    "Failed to Transition": "nepavyko perjungti",
    "KAM Power": "KAM atminties maitinimas",
    "Power Loss": "maitinimo praradimas",
    "KAM": "nuolatinė adaptyvioji atmintis (KAM)",
    "Warm, Downstream": "įšilęs, už katalizatoriaus esantis",
    "O2 Sensor": "deguonies (lambda) jutiklis",
    "Battery Temperature Sensor": "akumuliatoriaus temperatūros jutiklis",
    "Engine Temperature Sensor": "variklio temperatūros jutiklis",
    "Wheel Speed Sensor": "rato sukimosi greičio jutiklis",
    "Fan Speed Sensor": "ventiliatoriaus sukimosi greičio jutiklis",
    "Exhaust Pressure Sensor": "išmetamųjų dujų slėgio jutiklis",
    "Exhaust Pressure Control Valve": "išmetamųjų dujų slėgio valdymo vožtuvas",
    "Intake Valve Control": "įsiurbimo vožtuvo valdymas",
    "Exhaust Valve Control": "išmetimo vožtuvo valdymas",
    "Variable Intake Solenoid": "kintamojo įsiurbimo solenoidas",
    "Restraints System Sensor": "pasyviosios saugos sistemos jutiklis",
    "Lost Communication With Restraints System Sensor": "prarastas ryšys su pasyviosios saugos sistemos jutikliu",
    "MAF Sensor Intermittent": "protarpinis oro masės srauto jutiklio signalas",
    "Check of all OBDII Systems Not Complete": "nebaigta visų OBD-II sistemų patikra",
    "No Communication to TV Module": "nėra ryšio su TV moduliu",
    "No Communication to VICS Module": "nėra ryšio su VICS moduliu",
    "PATS Received Invalid Format Of Key-Code From Ignition Key Transponder": "PATS sistema iš uždegimo rakto atsakiklio gavo netinkamo formato rakto kodą",
    "Remote Keyless Entry": "nuotolinio berakčio patekimo į automobilį sistema",
    "Out of Synchronization": "sinchronizacijos neatitiktis",
    "Steering Wheel Angle": "vairo pasukimo kampas",
    "Steering EVO": "elektroninis vairo stiprintuvo valdymas (EVO)",
    "Throttle Control Unit": "droselinės sklendės valdymo blokas",
    "Loss of Return Spring": "grąžinimo spyruoklės praradimas",
    "Tire / Axle Ratio": "padangos / ašies perdavimo santykis",
    "Out of Acceptable Range": "už leistino diapazono ribų",
    "Time-Out": "skirtojo laiko viršijimas",
    "Traction Control of Brake": "stabdžių traukos kontrolė",
    "Traction Control of Engine": "variklio traukos kontrolė",
    "Window Feedback Loss of Signal": "elektra valdomo lango grįžtamojo ryšio signalo praradimas",
    "Signal Faulted": "signalo gedimas",
    "Correlation": "tarpusavio atitiktis",
    "Correlation Failure": "signalų tarpusavio neatitiktis",
    "Camshaft Position Actuator": "skirstomojo veleno padėties reguliatorius",
    "Crankshaft Position Camshaft Position Correlation": "alkūninio ir skirstomojo velenų padėties signalų tarpusavio neatitiktis",
    "Engine Oil Temp": "variklio alyvos temperatūra",
    "Engine Torque": "variklio sukimo momentas",
    "Function Read": "funkcijos duomenų nuskaitymas",
    "Audible Warnings": "garsiniai įspėjimai",
    "Anti-Theft Module": "apsaugos nuo vagystės sistemos modulis",
    "Vehicle Immobilized": "transporto priemonė imobilizuota",
    "Primary Id": "pagrindinis identifikatorius",
    "Throttle": "droselinė sklendė",
    "Traction Motor": "traukos elektros variklis",
    "Vehicle Speed Control": "transporto priemonės greičio valdymas",
    "Engine Oil Torque": "variklio sukimo momentas",
}

# Modulių pavadinimų galūnių sutvarkymas po frazės „Prarastas ryšys su“.
MODULE_SUFFIXES = {
    "variklio / jėgos pavaros valdymo blokas A": "variklio / jėgos pavaros valdymo bloku A",
    "variklio / jėgos pavaros valdymo blokas B": "variklio / jėgos pavaros valdymo bloku B",
    "variklio valdymo blokas": "variklio valdymo bloku",
    "pavarų dėžės valdymo blokas": "pavarų dėžės valdymo bloku",
    "stabdžių antiblokavimo sistema (ABS)": "stabdžių antiblokavimo sistemos (ABS) valdymo bloku",
}

_TRANSLATION_MAP = dict(TRANSLATIONS)
_TRANSLATION_MAP.update(PHRASE_TRANSLATIONS)
_TRANSLATION_MAP.update(EXTRA_PHRASE_TRANSLATIONS)
_COMPILED_TRANSLATIONS = [
    (re.compile(r"(?<![A-Za-z0-9])" + re.escape(source) + r"(?![A-Za-z0-9])", re.IGNORECASE), target)
    for source, target in sorted(_TRANSLATION_MAP.items(), key=lambda item: len(item[0]), reverse=True)
]

def replace_dictionary_phrases(value):
    """Keičia tik atskiras frazes / žodžius, niekada ne žodžio viduje."""
    for pattern, target in _COMPILED_TRANSLATIONS:
        value = pattern.sub(target, value)
    return re.sub(r"\s+", " ", value).strip()

def translate_description(english):
    """Verčia DTC aprašymus lokaliu techninių frazių žodynu, be interneto API."""
    text = re.sub(r"\s+", " ", str(english)).strip()
    if not text:
        return "Vertimas neprieinamas – tuščias aprašymas."

    exact_phrases = {
        "Exhaust Camshaft Position Actuator Circuit / Open": "Išmetimo skirstomojo veleno padėties reguliatoriaus elektros grandinė nutrūkusi",
        "Intake Camshaft Position Actuator Circuit / Open": "Įsiurbimo skirstomojo veleno padėties reguliatoriaus elektros grandinė nutrūkusi",
    }
    for source, target in exact_phrases.items():
        if text.lower().startswith(source.lower()):
            suffix = text[len(source):].strip()
            suffix = re.sub(r"(?i)^\(Bank\s*(\d+)\)$", lambda x: f"({x.group(1)}-oji cilindrų eilė)", suffix)
            return target + (" " + suffix if suffix else "") + "."

    if re.match(r"(?i)^STEERING Wheel Angle 1and 2 Circuit Failure$", text):
        return "Vairo pasukimo kampo jutiklių 1 ir 2 elektros grandinės gedimas."
    if re.match(r"(?i)^STEERING Wheel Angle 1and 2 Signal Faulted$", text):
        return "Vairo pasukimo kampo jutiklių 1 ir 2 signalų gedimas."
    if re.match(r"(?i)^Steering EVO Out-of-Range Fault$", text):
        return "Elektroninio vairo stiprintuvo valdymo (EVO) signalo reikšmė nepatenka į leistiną diapazoną."
    if re.match(r"(?i)^Evaporative Emission Control System Pressure Sensor Low Input$", text):
        return "Degalų garų kontrolės sistemos (EVAP) slėgio jutiklio įėjimo signalas per žemas."
    if re.match(r"(?i)^ABS Inlet Valve Coil Rear Circuit Short To Ground$", text):
        return "ABS galinio rato įleidimo vožtuvo ritės elektros grandinės trumpasis jungimas į masę."

    # Pirmiausia apdorojamos pilnos OBD diagnostikos konstrukcijos.
    # Tai apsaugo nuo pažodinio atskirų žodžių vertimo ir netaisyklingos žodžių tvarkos.
    def translate_component(value):
        value = value.strip(" -/:;,. ")
        value = replace_dictionary_phrases(value)
        return value[:1].lower() + value[1:] if value else value

    # „X Control Circuit ...“ – dažna P/B/C serijų konstrukcija.
    m = re.match(r"(?i)^(.+?)\s+Control Circuit\s*(?:/\s*Open|Open Circuit)$", text)
    if m:
        return f"{translate_component(m.group(1))} valdymo elektros grandinė nutrūkusi."
    m = re.match(r"(?i)^(.+?)\s+Control Circuit\s+(Low|High)(?:\s+Input)?$", text)
    if m:
        level = "per žemas" if m.group(2).lower() == "low" else "per aukštas"
        return f"{translate_component(m.group(1))} valdymo grandinės signalas {level}."
    m = re.match(r"(?i)^(.+?)\s+Control Circuit\s+(?:Range/Performance|Range / Performance)$", text)
    if m:
        return f"{translate_component(m.group(1))} valdymo grandinės veikimo diapazono / našumo neatitiktis."
    m = re.match(r"(?i)^(.+?)\s+Circuit\s+(Low|High)(?:\s+Input)?$", text)
    if m:
        level = "per žemas" if m.group(2).lower() == "low" else "per aukštas"
        return f"{translate_component(m.group(1))} elektros grandinės signalas {level}."
    m = re.match(r"(?i)^(.+?)\s+Circuit\s+(?:/\s*)?Open$", text)
    if m:
        return f"{translate_component(m.group(1))} elektros grandinė nutrūkusi."
    m = re.match(r"(?i)^(.+?)\s+Circuit\s+(?:Range/Performance|Range / Performance)$", text)
    if m:
        return f"{translate_component(m.group(1))} elektros grandinės veikimo diapazono / našumo neatitiktis."
    m = re.match(r"(?i)^(.+?)\s+Circuit\s+(?:Malfunction|Failure)$", text)
    if m:
        return f"{translate_component(m.group(1))} elektros grandinės gedimas."
    m = re.match(r"(?i)^(.+?)\s+Position\s+Timing\s*-?\s*(Over-Advanced|Over-Retarded)(?:\s*\(Bank\s*(\d+)\))?$", text)
    if m:
        timing = "pernelyg paankstintas" if m.group(2).lower() == "over-advanced" else "pernelyg suvėlintas"
        bank = f" ({m.group(3)}-oji cilindrų eilė)" if m.group(3) else ""
        return f"{translate_component(m.group(1))} fazė {timing}{bank}."
    m = re.match(r"(?i)^Crankshaft Position Camshaft Position Correlation Bank\s*(\d+) Sensor\s*([A-Z0-9]+)$", text)
    if m:
        return f"Alkūninio ir skirstomojo velenų padėties signalų tarpusavio neatitiktis ({m.group(1)}-oji cilindrų eilė, jutiklis {m.group(2)})."

    # U serija turi savitas tinklo diagnostikos formuluotes. Jas verčiame
    # atskiromis taisyklėmis, kad neliktų angliškų modulių pavadinimų.
    module_map = {
      "Engine Control Module":"variklio valdymo modulis",
      "Powertrain Control Module":"jėgos pavaros valdymo modulis",
      "Transmission Control Module":"pavarų dėžės valdymo modulis",
      "Anti-Lock Brake System Control Module":"stabdžių antiblokavimo sistemos (ABS) valdymo modulis",
      "ABS Control Module":"stabdžių antiblokavimo sistemos (ABS) valdymo modulis",
      "Instrument Panel Control Module":"prietaisų skydelio valdymo modulis",
      "Instrument Panel Cluster":"prietaisų skydelio blokas",
      "Body Control Module":"kėbulo elektronikos valdymo modulis",
      "Electronic Brake Control Module":"elektroninis stabdžių valdymo modulis",
      "Steering Control Module":"vairo sistemos valdymo modulis",
      "Airbag Control Module":"oro pagalvių valdymo modulis",
      "Restraints Control Module":"pasyviosios saugos sistemos valdymo modulis",
      "Restraints System Sensor A":"pasyviosios saugos sistemos jutiklis A",
      "Restraints System Sensor B":"pasyviosios saugos sistemos jutiklis B",
      "Restraints System Sensor C":"pasyviosios saugos sistemos jutiklis C",
      "Restraints System Sensor D":"pasyviosios saugos sistemos jutiklis D",
      "Restraints System Sensor E":"pasyviosios saugos sistemos jutiklis E",
      "Restraints System Sensor F":"pasyviosios saugos sistemos jutiklis F",
      "Restraints System Sensor G":"pasyviosios saugos sistemos jutiklis G",
      "Restraints System Sensor H":"pasyviosios saugos sistemos jutiklis H",
      "Restraints System Sensor I":"pasyviosios saugos sistemos jutiklis I",
      "Restraints System Sensor J":"pasyviosios saugos sistemos jutiklis J",
      "Gateway Control Module":"tinklo sąsajos (Gateway) valdymo modulis",
      "Central Gateway Module":"centrinis tinklo sąsajos modulis",
      "Fuel Pump Control Module":"degalų siurblio valdymo modulis",
      "Glow Plug Control Module":"kaitinimo žvakių valdymo modulis",
      "Parking Assist Control Module":"parkavimo pagalbos sistemos valdymo modulis",
      "Climate Control Module":"klimato kontrolės sistemos valdymo modulis",
      "HVAC Control Module":"šildymo, vėdinimo ir oro kondicionavimo sistemos valdymo modulis",
      "Active Roll Control Module":"aktyviosios kėbulo posvyrio kontrolės valdymo modulis",
      "All Terrain Control Module":"važiavimo bekele sistemos valdymo modulis",
      "Alternative Fuel Control Module":"alternatyviųjų degalų sistemos valdymo modulis",
      "Antenna Control Module":"antenos valdymo modulis",
      "Anti-Lock Brake System (ABS) Control Module":"stabdžių antiblokavimo sistemos (ABS) valdymo modulis",
      "Audible Alert Control Module":"garsinių įspėjimų valdymo modulis",
      "Audio Amplifier A":"garso stiprintuvas A",
      "Automatic Lighting Control Module":"automatinio apšvietimo valdymo modulis",
      "Auxiliary Heater Control Module":"papildomo šildytuvo valdymo modulis",
      "Battery Energy Control Module A":"akumuliatoriaus energijos valdymo modulis A",
      "Battery Energy Control Module B":"akumuliatoriaus energijos valdymo modulis B",
      "Body Control Module":"kėbulo elektronikos valdymo modulis",
      "Body Control Module A":"kėbulo elektronikos valdymo modulis A",
      "Body Control Module B":"kėbulo elektronikos valdymo modulis B",
      "Body Control Module C":"kėbulo elektronikos valdymo modulis C",
      "Body Control Module D":"kėbulo elektronikos valdymo modulis D",
      "Body Control Module E":"kėbulo elektronikos valdymo modulis E",
      "Brake System Control Module":"stabdžių sistemos valdymo modulis",
      "Compass Module":"kompaso modulis",
      "Coolant Temperature Control Module":"aušinimo skysčio temperatūros valdymo modulis",
      "Cruise Control Module":"pastovaus greičio palaikymo sistemos valdymo modulis",
      "Differential Control ModuleFront":"priekinio diferencialo valdymo modulis",
      "Differential Control ModuleRear":"galinio diferencialo valdymo modulis",
      "Digital Audio Control Module A":"skaitmeninės garso sistemos valdymo modulis A",
      "Digital Audio Control Module B":"skaitmeninės garso sistemos valdymo modulis B",
      "Digital Disc Player/Changer Module A":"skaitmeninių diskų grotuvo / keitiklio modulis A",
      "Digital Disc Player/Changer Module B":"skaitmeninių diskų grotuvo / keitiklio modulis B",
      "Digital Disc Player/Changer Module C":"skaitmeninių diskų grotuvo / keitiklio modulis C",
      "Digital Disc Player/Changer Module D":"skaitmeninių diskų grotuvo / keitiklio modulis D",
      "Door Control Module A":"durelių valdymo modulis A",
      "Drive Motor Control Module":"traukos variklio valdymo modulis",
      "ECM/PCM A":"variklio / jėgos pavaros valdymo modulis A",
      "ECM/PCM B":"variklio / jėgos pavaros valdymo modulis B",
      "Emissions Critical Control Information":"su išmetamųjų teršalų kontrole susijusi informacija",
      "Four-Wheel Drive Clutch Control Module":"visų varančiųjų ratų sankabos valdymo modulis",
      "Front Lighting Control Module":"priekinio apšvietimo valdymo modulis",
      "Fuel Additive Control Module":"degalų priedo valdymo modulis",
      "Fuel Cell Control Module":"kuro elemento valdymo modulis",
      "Fuel Injector Control Module":"degalų purkštukų valdymo modulis",
      "Fuel Pump Control Module":"degalų siurblio valdymo modulis",
      "Gateway A":"tinklo sąsajos modulis A", "Gateway B":"tinklo sąsajos modulis B",
      "Gateway C":"tinklo sąsajos modulis C", "Gateway D":"tinklo sąsajos modulis D", "Gateway E":"tinklo sąsajos modulis E",
      "Gear Shift Module":"pavarų perjungimo modulis",
      "Glow Plug Control Module":"kaitinimo žvakių valdymo modulis",
      "HVAC Control Module":"šildymo, vėdinimo ir oro kondicionavimo sistemos valdymo modulis",
      "Head Up Display":"projekcinis ekranas",
      "Headlamp Leveling Control Module":"žibintų aukščio reguliavimo valdymo modulis",
      "Information Center A":"informacinis blokas A", "Information Center B":"informacinis blokas B",
      "Instrument Panel Cluster (IPC) Control Module":"prietaisų skydelio (IPC) valdymo modulis",
      "Lateral Acceleration Sensor Module":"šoninio pagreičio jutiklio modulis",
      "Multi-axis Acceleration Sensor Module":"daugi ašių pagreičio jutiklio modulis",
      "Navigation Control Module":"navigacijos sistemos valdymo modulis",
      "Navigation Display Module":"navigacijos ekrano modulis",
      "PTO Control Module":"galios nuėmimo įrenginio valdymo modulis",
      "Park Brake Control Module":"stovėjimo stabdžio valdymo modulis",
      "Parking Assist Control Module A":"parkavimo pagalbos sistemos valdymo modulis A",
      "Personal Computer":"asmeninis kompiuteris",
      "Power Steering Control Module":"vairo stiprintuvo valdymo modulis",
      "Power Steering Control Module Rear":"galinės vairo stiprinimo sistemos valdymo modulis",
      "Radio":"radijo imtuvas",
      "Rear HVAC Control Module":"galinės klimato kontrolės sistemos valdymo modulis",
      "Rear Lighting Control Module":"galinio apšvietimo valdymo modulis",
      "Rear Seat Entertainment Control Module":"galinių sėdynių pramogų sistemos valdymo modulis",
      "Restraints Control Module":"pasyviosios saugos sistemos valdymo modulis",
      "Restraints System Sensor A":"pasyviosios saugos sistemos jutiklis A",
      "Restraints System Sensor B":"pasyviosios saugos sistemos jutiklis B",
      "Restraints System Sensor C":"pasyviosios saugos sistemos jutiklis C",
      "Restraints System Sensor D":"pasyviosios saugos sistemos jutiklis D",
      "Restraints System Sensor E":"pasyviosios saugos sistemos jutiklis E",
      "Restraints System Sensor F":"pasyviosios saugos sistemos jutiklis F",
      "Restraints System Sensor G":"pasyviosios saugos sistemos jutiklis G",
      "Restraints System Sensor H":"pasyviosios saugos sistemos jutiklis H",
      "Restraints System Sensor I":"pasyviosios saugos sistemos jutiklis I",
      "Restraints System Sensor J":"pasyviosios saugos sistemos jutiklis J",
      "Restraints Occupant Sensing Control Module":"keleivių atpažinimo sistemos valdymo modulis",
      "Ride Level Control Module":"kėbulo aukščio reguliavimo valdymo modulis",
      "Side Restraints Control Module Left":"kairės pusės šoninės saugos sistemos valdymo modulis",
      "Side Restraints Control Module Right":"dešinės pusės šoninės saugos sistemos valdymo modulis",
      "Starter / Generator Control Module":"starterio ir generatoriaus valdymo modulis",
      "Steering Angle Sensor Module":"vairo kampo jutiklio modulis",
      "Steering Effort Control Module":"vairo pasipriešinimo valdymo modulis",
      "Subscription Entertainment Receiver Module":"prenumeruojamų pramogų paslaugų imtuvo modulis",
      "Sunroof Control Module":"stoglangio valdymo modulis",
      "Suspension Control Module B":"pakabos valdymo modulis B",
      "Telematic Control Module":"telematikos valdymo modulis",
      "Telephone Control Module":"telefono sistemos valdymo modulis",
      "Television":"televizorius",
      "Throttle Actuator Control Module":"droselinės sklendės pavaros valdymo modulis",
      "Tire Pressure Monitor Module":"padangų slėgio stebėjimo sistemos valdymo modulis",
      "Trailer Brake Control Module":"priekabos stabdžių valdymo modulis",
      "Vehicle Dynamics Control Module":"automobilio dinamikos valdymo modulis",
      "Vehicle Immobilizer Control Module":"imobilizatoriaus valdymo modulis",
      "Vehicle Security Control Module":"automobilio apsaugos sistemos valdymo modulis",
      "Yaw Rate Sensor Module":"sukimosi apie vertikaliąją ašį greičio jutiklio modulis",
      "TCM":"pavarų dėžės valdymo modulis (TCM)",
      "Transfer Case Control Module":"skirstomosios pavarų dėžės valdymo modulis",
    }
    lost = re.match(r"(?i)^Lost Communication With\s+(.+)$", text)
    if lost:
        name = lost.group(1).strip().rstrip(".")
        translated = module_map.get(name)
        if translated:
            return f"Prarastas ryšys su moduliu: {translated}."
        # Nežinomam moduliui paliekamas originalus pavadinimas, aiškiai nurodant, kad reikia patikslinti.
        return f"Prarastas ryšys su valdymo moduliu: {name}. Modulio pavadinimo vertimą reikia patikslinti pagal gamintojo dokumentaciją."

    if re.match(r"(?i)^No Communication to TV Module \(No Fitting of TV\)$", text):
        return "Nėra ryšio su TV moduliu (šis modulis transporto priemonėje neįrengtas)."
    if re.match(r"(?i)^No Communication to VICS Module \(No Fitting of VICS\)$", text):
        return "Nėra ryšio su VICS moduliu (šis modulis transporto priemonėje neįrengtas)."

    # Kiti dažni U serijos įrašai (ne tik „Lost Communication“).
    m = re.match(r"(?i)^(.+?)\s+is not responding$", text)
    if m:
        subject = m.group(1).strip()
        names = {"Audio Bezel":"garso sistemos valdymo skydelis", "Audio Center Amp":"centrinis garso stiprintuvas",
          "Audio Compact Disk / Disk Jockey Unit":"CD grotuvas / diskų keitiklis", "Audio Front Control Module (ACM)":"priekinis garso sistemos valdymo modulis (ACM)",
          "Audio Navigation Unit":"garso ir navigacijos blokas", "Audio Phone":"automobilio telefono sistema", "Audio Rear Control Unit":"galinis garso valdymo blokas",
          "Audio Rear Integrated Control Panel Unit":"galinis integruotas valdymo skydelis", "Audio Remote Climate Control Unit":"nuotolinis klimato kontrolės blokas",
          "Audio Steering Wheel Control Unit":"garso sistemos valdymo ant vairo blokas", "Audio Subwoofer Unit":"žemųjų dažnių garsiakalbio blokas",
          "Audio Tape Deck Unit":"kasetinis grotuvas", "Audio Voice Module":"balso valdymo modulis", "Compass Module":"kompaso modulis"}
        return "Neatsako: " + names.get(subject, subject) + "."

    m = re.match(r"(?i)^SCP \(J1850\) Invalid or Missing Data for (.+)$", text)
    if m:
        topics = {"Digital Signal Processing":"skaitmeninį signalų apdorojimą", "Engine Oil Temp":"variklio alyvos temperatūrą", "Engine Torque":"variklio sukimo momentą", "Function Read Audible Warnings / Anti-Theft Module, Vehicle Immobilized":"garsinių įspėjimų ir apsaugos nuo vagystės sistemos modulio funkciją, kai transporto priemonė imobilizuota", "Function Read Engine Torque":"variklio sukimo momento nuskaitymo funkciją", "Function Read Fuel System":"degalų sistemos nuskaitymo funkciją", "Function Read Vehicle Speed":"transporto priemonės greičio nuskaitymo funkciją", "Primary Id / Vehicle Speed":"pagrindinį identifikatorių / transporto priemonės greitį", "Throttle":"droselinę sklendę", "Traction Motor":"traukos elektros variklį", "Vehicle Speed Control":"transporto priemonės greičio valdymą", "Brakes":"stabdžių sistemą", "Engine Coolant":"variklio aušinimo skystį", "Engine Oil":"variklio alyvą", "Engine RPM":"variklio sūkių dažnį",
          "Fuel System":"degalų sistemą", "Vehicle Speed":"transporto priemonės greitį", "Traction Control":"traukos kontrolės sistemą",
          "Transmission / Transaxle / PRNDL":"pavarų dėžę / transmisiją / pavarų pasirinkimo indikatorių", "Vehicle Id (VIN)":"transporto priemonės identifikavimo numerį (VIN)",
          "Climate Control (HVAC)":"klimato kontrolės sistemą (HVAC)", "Air Conditioning Clutch":"oro kondicionieriaus sankabą", "Door Locks":"durelių užraktus",
          "Steering / Steering Wheel":"vairo sistemą / vairą", "Windows":"elektra valdomus langus", "Wheels":"ratus", "Suspension":"pakabą", "Primary Id":"pagrindinį identifikatorių",
          "Engine Sensors":"variklio jutiklius", "Engine Systems Other":"kitas variklio sistemas", "Powertrain Status Request":"užklausą apie jėgos pavaros būseną",
          "Chassis Status Request":"užklausą apie važiuoklės būseną", "Body Status Request":"užklausą apie kėbulo sistemų būseną",
          "Non-Legislated Diagnostics":"teisės aktais nereglamentuojamą diagnostiką", "Legislated Diagnostics":"teisės aktais reglamentuojamą diagnostiką",
          "Electric Traction Drive (Inverter)":"elektrinę traukos pavarą (keitiklį)", "Electrical Energy Management":"elektros energijos valdymą",
          "Odometer":"odometrą", "Vehicle Motion":"transporto priemonės judėjimą", "Telltales":"įspėjamąsias indikacijas",
          "Audio Control":"garso sistemą", "Audible Warnings":"garsinius įspėjimus", "Compact Disc":"kompaktinį diską",
          "Digital Signal Processing":"skaitmeninį signalų apdorojimą", "Digital Audio Tape":"skaitmeninę garso juostą",
          "Tuner / Receiver":"radijo imtuvą", "Cassette Tape":"kasetinę juostą", "Cellular Phone / Paging System":"mobiliojo telefono / pranešimų gaviklio sistemą",
          "Remote Button Control":"nuotolinio mygtukų valdymo sistemą", "Personalization (Memory) Features":"individualių nustatymų (atminties) funkcijas",
          "Window Wiper / Washer":"langų valytuvų / plovimo sistemą", "Seat Motion / Control":"sėdynės judėjimo / valdymo sistemą",
          "Steering Column":"vairo kolonėlę", "Seat Switches":"sėdynės jungiklius", "External Lamp Outage":"išorinių žibintų gedimą",
          "External Lamps":"išorinius žibintus", "Interior Lamp Outage":"salono apšvietimo gedimą", "Interior Lamps":"salono apšvietimo žibintus",
          "Tires":"padangas", "Electric Defrost":"elektrinį stiklo šildymą", "Displays":"ekranus", "Memory Storage":"atminties įrenginį",
          "Exterior Environment":"išorinės aplinkos duomenis", "Interior Environment":"salono aplinkos duomenis", "Time / Date":"laiką / datą",
          "Class A Functions":"A klasės funkcijas", "Network Control":"tinklo valdymą", "Experimental #1":"eksperimentinį parametrą Nr. 1",
          "Experimental #2":"eksperimentinį parametrą Nr. 2", "Experimental #3":"eksperimentinį parametrą Nr. 3", "Experimental #4":"eksperimentinį parametrą Nr. 4",
          "Vehicle Configuration":"transporto priemonės konfigūraciją", "Vehicle Security":"automobilio apsaugos sistemą", "Engine Air Intake":"variklio oro įsiurbimo sistemą",
          "Charging System":"įkrovimo sistemą", "Ignition Switch / Starter":"uždegimo jungiklį / starterį", "Mirrors":"veidrodėlius", "Restraints":"pasyviosios saugos sistemą"}
        topic=m.group(1).strip(); return "SCP (J1850) magistralėje gaunami neteisingi duomenys arba jų trūksta apie " + topics.get(topic, topic.lower()) + "."

    m = re.match(r"(?i)^SCP \(J1850\) Invalid Data from (.+)$", text)
    if m:
        blocks={"GEM":"bendrąjį elektroninį modulį (GEM)","IC":"prietaisų skydelio bloką (IC)","REM":"galinį elektronikos modulį (REM)","SCLM":"vairo kolonėlės užrakto modulį (SCLM)"}
        return "SCP (J1850) magistralėje gauti neteisingi duomenys iš " + ({"GEM":"bendrojo elektroninio modulio (GEM)","IC":"prietaisų skydelio bloko (IC)","REM":"galinio elektronikos modulio (REM)","SCLM":"vairo kolonėlės užrakto modulio (SCLM)"}.get(m.group(1).strip(),m.group(1))) + "."

    m = re.match(r"(?i)^SCP \(J1850\) Lack of Acknowledgment for (.+)$", text)
    if m:
        terms={"Audible Warnings":"garsinius įspėjimus", "Door Locks":"durelių užraktus", "External Access (Doors)":"išorinę prieigą (dureles)", "Mirrors":"veidrodėlius", "Primary Id":"pagrindinį identifikatorių", "Telltales":"įspėjamąsias indikacijas"}
        return "SCP (J1850) magistralėje negautas patvirtinimas dėl " + terms.get(m.group(1).strip(),m.group(1).lower()) + "."

    if re.match(r"(?i)^High Speed CAN Communication Bus$", text): return "Didelės spartos CAN ryšio magistralė."
    if re.match(r"(?i)^CAN Communication Bus Fault$", text): return "CAN ryšio magistralės gedimas."
    if re.match(r"(?i)^UPB Communication Bus Fault$", text): return "UPB ryšio magistralės gedimas."
    if re.match(r"(?i)^SCP \(J1850\) Communication Bus Fault$", text): return "SCP (J1850) ryšio magistralės gedimas."
    if re.match(r"(?i)^Communication Bus Error \(Non SCP\)$", text): return "Ryšio magistralės klaida (ne SCP sistema)."
    if re.match(r"(?i)^\(CAN\) Lack of Acknowledgement From Engine Management$", text): return "CAN magistralėje negautas variklio valdymo sistemos patvirtinimas."
    if re.match(r"(?i)^Module is Not Responding \(Non SCP\)$", text): return "Valdymo modulis neatsako (ne SCP sistema)."
    if re.match(r"(?i)^Module Transmitted Invalid Data \(Non SCP\)$", text): return "Valdymo modulis perdavė neteisingus duomenis (ne SCP sistema)."
    if re.match(r"(?i)^Invalid /?fault data received \(Non SCP\)$", text): return "Gauti neteisingi arba klaidingi duomenys (ne SCP sistema)."
    if re.match(r"(?i)^Signal Link Fault \(Non SCP\)$", text): return "Signalo ryšio linijos gedimas (ne SCP sistema)."
    if re.match(r"(?i)^Signal Link Short to Ground \(Non SCP\) / From NGV Module$", text): return "Signalo ryšio linijos trumpasis jungimas į masę (ne SCP sistema) / iš NGV modulio."
    if re.match(r"(?i)^Driver Side Crash Sensor Communication Fault \(Non SCP\)$", text): return "Vairuotojo pusės smūgio jutiklio ryšio gedimas (ne SCP sistema)."
    if re.match(r"(?i)^Passenger Side Crash Sensor Communication Fault \(Non SCP\)$", text): return "Keleivio pusės smūgio jutiklio ryšio gedimas (ne SCP sistema)."

    m = re.match(r"(?i)^STEERING Wheel Angle 1and 2 Circuit Failure$", text)
    if m:
        return "Vairo pasukimo kampo jutiklių 1 ir 2 elektros grandinės gedimas."
    m = re.match(r"(?i)^STEERING Wheel Angle 1and 2 Signal Faulted$", text)
    if m:
        return "Vairo pasukimo kampo jutiklių 1 ir 2 signalų gedimas."
    m = re.match(r"(?i)^Steering EVO Out-of-Range Fault$", text)
    if m:
        return "Elektroninio vairo stiprintuvo valdymo (EVO) signalo reikšmė nepatenka į leistiną diapazoną."

    # Frazės keičiamos nuo ilgiausių iki trumpiausių, kad bendri žodžiai
    # nesugadintų jau išversto techninio termino.
    text = replace_dictionary_phrases(text)

    # Papildomos santrumpos, skaičiai ir rašybos normalizavimas.
    text = re.sub(r"\bHO2S\b|\bH02S\b", "lambda jutiklis", text, flags=re.IGNORECASE)
    text = re.sub(r"\bMAF\b", "oro masės srauto jutiklis (MAF)", text, flags=re.IGNORECASE)
    text = re.sub(r"\bMAP\b", "įsiurbimo kolektoriaus slėgio jutiklis (MAP)", text, flags=re.IGNORECASE)
    text = re.sub(r"\bEIC\b", "elektroninis prietaisų skydelio valdiklis (EIC)", text, flags=re.IGNORECASE)
    text = re.sub(r"\bckt\b", "grandinė", text, flags=re.IGNORECASE)
    text = re.sub(r"\bBank\s*(\d+)\b", lambda m: f"{m.group(1)}-oji cilindrų eilė", text, flags=re.IGNORECASE)
    text = re.sub(r"\bSensor\s*(\d+)\b", lambda m: f"jutiklis {m.group(1)}", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"([,;])(?=\S)", r"\1 ", text)
    text = text.strip(" -/.")
    if not text:
        return "Vertimas neprieinamas."
    return text[0].upper() + text[1:] + ("." if text[-1] not in ".!?" else "")

def manufacturer_status(code):
    """Grąžina Ford / gamintojo specifikos pranešimą, nepervertindama tikrumo."""
    code = code.upper()
    if code in FORD_SPECIFIC_CODES:
        return (
            "FORD SPECIFINIS KODAS",
            "Šis kodas įtrauktas į programos Ford specifinių kodų sąrašą. "
            "Patikrinkite, ar aprašymas atitinka konkretaus automobilio modelį, metus ir variklį."
        )

    # P1xxx ir P3xxx bei atitinkami kėbulo, važiuoklės ir ryšio kodų diapazonai
    # yra gamintojo valdomi; vien prefiksas neįrodo, kad kodas yra Ford.
    if re.fullmatch(r"P[13]\d{3}|[BCU]1\d{3}", code):
        return (
            "GALIMAI GAMINTOJO SPECIFINIS KODAS",
            "Šis kodas yra gamintojo valdomame OBD-II diapazone. Jis gali būti "
            "Ford specifinis, tačiau vien pagal kodo formatą to patvirtinti negalima. "
            "Reikalinga Ford dokumentacija pagal modelį, pagaminimo metus ir variklį."
        )

    return (
        "STANDARTINIS OBD-II DIAPAZONAS",
        "Pagal kodo prefiksą tai standartizuoto OBD-II diapazono kodas. "
        "Vis tiek gali būti gamintojo papildomų interpretavimo ypatumų."
    )

def possible_faults(code, description):
    """Pateikia orientacines priežastis pagal aprašymo raktažodžius."""
    d = description.lower()
    faults = []

    if any(k in d for k in ["oxygen sensor", "ho2s", "o2 sensor", "lambda"]):
        faults = [
            "Sugedęs lambda (deguonies) jutiklis arba jo kaitinimo elementas.",
            "Pažeisti jutiklio laidai, jungtis, maitinimas arba masė.",
            "Išmetimo sistemos nesandarumas arba netinkamas jutiklio signalas."
        ]
    elif any(k in d for k in ["mass or volume air", "maf", "map", "manifold absolute pressure", "boost sensor"]):
        faults = [
            "Užterštas arba sugedęs oro srauto / slėgio jutiklis.",
            "Pažeista jutiklio elektros instaliacija ar jungtis.",
            "Įsiurbimo sistemos nesandarumas, užsikimšęs filtras arba netinkamas slėgis."
        ]
    elif any(k in d for k in ["injector", "injection", "fuel rail", "fuel pressure", "fuel pump", "fuel delivery"]):
        faults = [
            "Purkštuko, degalų siurblio arba slėgio reguliatoriaus gedimas.",
            "Užsikimšęs degalų filtras arba ribotas degalų tiekimas.",
            "Laidų, jungčių, maitinimo ar masės grandinės pažeidimas."
        ]
    elif any(k in d for k in ["glow plug", "heater circuit"]):
        faults = [
            "Perdegusi kaitinimo žvakė arba sugedęs kaitinimo elementas.",
            "Kaitinimo žvakių relės / valdymo modulio gedimas.",
            "Pažeisti laidai, saugiklis arba jungtys."
        ]
    elif any(k in d for k in ["turbo", "boost", "wastegate"]):
        faults = [
            "Turbokompresoriaus valdymo vožtuvo arba aktuatoriaus gedimas.",
            "Vakuumo žarnelių, slėgio žarnų ar intercoolerio nesandarumas.",
            "Užstrigusi turbinos geometrija arba neteisingas slėgio jutiklio signalas."
        ]
    elif any(k in d for k in ["egr", "exhaust gas recirculation"]):
        faults = [
            "Užsikimšęs arba užstrigęs EGR vožtuvas.",
            "EGR valdymo vožtuvo, padėties jutiklio ar vakuumo sistemos gedimas.",
            "Pažeisti laidai, jungtys arba užteršti EGR kanalai."
        ]
    elif any(k in d for k in ["coolant", "thermostat", "cooling fan", "engine overtemp"]):
        faults = [
            "Aušinimo skysčio temperatūros jutiklio arba termostato gedimas.",
            "Per mažas aušinimo skysčio kiekis, nuotėkis arba oro kamštis sistemoje.",
            "Ventiliatoriaus, relės, saugiklio ar laidų gedimas."
        ]
    elif any(k in d for k in ["crankshaft", "camshaft", "engine position", "timing"]):
        faults = [
            "Alkūninio arba skirstomojo veleno padėties jutiklio gedimas.",
            "Pažeisti jutiklio laidai, jungtys arba maitinimo grandinė.",
            "Netinkamas paskirstymo mechanizmo fazių sureguliavimas."
        ]
    elif any(k in d for k in ["misfire", "contribution/balance"]):
        faults = [
            "Uždegimo sistemos arba kaitinimo / degimo proceso sutrikimas.",
            "Purkštuko gedimas, netolygus degalų tiekimas arba per mažas suspaudimas.",
            "Oro patekimas į įsiurbimo sistemą arba mechaninis variklio gedimas."
        ]
    elif any(k in d for k in ["voltage", "circuit", "signal", "electrical", "relay"]):
        faults = [
            "Pažeisti laidai, izoliacija, jungtys arba kontaktai.",
            "Perdegęs saugiklis, sugedusi relė arba netinkamas maitinimas / masė.",
            "Sugedęs atitinkamas jutiklis, vykdiklis arba valdymo blokas."
        ]
    elif any(k in d for k in ["control module", "processor", "memory", "ecm", "pcm", "tcm"]):
        faults = [
            "Valdymo bloko maitinimo arba masės grandinės sutrikimas.",
            "Pažeista elektros instaliacija arba ryšio grandinė.",
            "Valdymo bloko programinės įrangos ar vidinis gedimas (patvirtinti tik atlikus patikrą)."
        ]
    else:
        faults = [
            "Galimas atitinkamo jutiklio, vykdiklio arba mechaninio mazgo gedimas.",
            "Galimas laidų, jungčių, maitinimo arba masės grandinės sutrikimas.",
            "Būtina patikrinti gamintojo diagnostikos procedūrą ir realius parametrus."
        ]

    return faults

try:
    DTC_DATABASE = extract_codes(DOCX_PATH)
    LOAD_ERROR = None
    print(f"Duomenų bazė įkelta. Rasta {len(DTC_DATABASE)} DTC kodų.")
except Exception as error:
    DTC_DATABASE = {}
    LOAD_ERROR = str(error)
    print("KLAIDA įkeliant duomenų bazę:")
    print(error)

INVALID_FORMAT_MESSAGE = 'Neteisingas formatas. Įveskite, pavyzdžiui, P0101, B1234, C1234 arba U0100.'
NOT_FOUND_DETAIL = 'Tai nereiškia, kad kodo nėra automobilio sistemoje – jis gali būti nepateiktame šaltinyje.'
DISCLAIMER = 'Pastaba: tai galimos priežastys, o ne galutinė diagnozė. Prieš keičiant dalis reikia atlikti diagnostiką ir patikrinti gamintojo techninius duomenis.'


def lookup_dtc(raw_code):
    # Same fields the notebook printed, returned for the web interface.
    code = str(raw_code).strip().upper().replace(" ", "")
    if not re.fullmatch(r"[PBCU]\d{4}", code):
        return {
            "ok": False,
            "error": "invalid_format",
            "code": code,
            "message": INVALID_FORMAT_MESSAGE,
        }

    if code not in DTC_DATABASE:
        return {
            "ok": False,
            "error": "not_found",
            "code": code,
            "message": f"Kodas {code} Word faile nerastas.",
            "detail": NOT_FOUND_DETAIL,
        }

    description = DTC_DATABASE[code]
    status, status_info = manufacturer_status(code)
    result = {
        "ok": True,
        "code": code,
        "description": description,
        "translation": translate_description(description),
        "status": status,
        "status_info": status_info,
        "faults": possible_faults(code, description),
        "disclaimer": DISCLAIMER,
    }
    if code in FORD_SPECIFIC_CODES:
        result["ford_note"] = FORD_SPECIFIC_CODES[code]
    return result


def suggest_codes(raw_query, limit=8):
    # Prefix matches for the search field. Full lookup stays exact.
    query = str(raw_query).strip().upper().replace(" ", "")
    if not query:
        return []
    found = []
    for code in sorted(DTC_DATABASE):
        if not code.startswith(query):
            continue
        description = DTC_DATABASE[code]
        if len(description) > 88:
            description = description[:87].rstrip() + "…"
        found.append({"code": code, "description": description})
        if len(found) >= limit:
            break
    return found

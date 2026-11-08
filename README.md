# Audilekt

Web-Dienst fürs Heimnetz: DOCX oder TXT hochladen, MP3 zurückbekommen. Die Sprachsynthese läuft lokal auf der eigenen GPU, die Oberfläche ist auch fürs Smartphone gemacht.

**Built with Fish Audio.** Audilekt basiert auf [Fish Audio Drama 3](https://huggingface.co/fishaudio/drama-3) und wird über [SGLang-Omni](https://github.com/sgl-project/sglang-omni) ausgeführt.

## Funktionen

- DOCX und TXT annehmen, mehrere Dateien nacheinander abarbeiten
- Text in Abschnitte teilen und parallel synthetisieren
- Stille und Endlosschleifen je Abschnitt erkennen, neu erzeugen und unsichere Stellen anzeigen
- Nach Abbruch oder Neustart beim letzten fertigen Abschnitt weitermachen
- Das Modell beim ersten Auftrag laden und nach Leerlauf wieder aus dem Grafikspeicher nehmen
- Mehrere Stimmen über kurze Referenzclips

## Voraussetzungen

- Linux x86_64 mit NVIDIA-GPU und mindestens 48 GB VRAM. Die Voreinstellungen sind auf 96 GB ausgelegt, entwickelt wurde auf einer RTX PRO 6000 Blackwell.
- NVIDIA-Treiber für CUDA 13
- Conda
- `ffmpeg` mit `libmp3lame`
- Rund 42 GB Speicherplatz für das Modell und 12 GB für die Python-Umgebung

## Installation

```bash
git clone https://github.com/Dimitri-Ingenieur/audilekt.git
cd audilekt

conda create -n audilekt python=3.12 -y
conda activate audilekt
pip install --no-deps -r requirements.txt

hf download fishaudio/drama-3 --local-dir ./drama-3

python tts_webui.py
```

Die Seite läuft danach unter `http://<server-ip>:8090`.

`--no-deps` ist Pflicht. Die Datei ist vollständig gepinnt, und `descript-audiotools` verlangt ein `protobuf`, das mit SGLang kollidiert. Ohne das Flag bricht pip ab.

### Modell

Drama 3 gehört nicht zu diesem Repository und muss von Hugging Face geladen werden (41,5 GB). Der Ordner gehört ins Projektverzeichnis, dort sucht `tts_webui.py` auch nach `codec.pth`:

```text
audilekt/
├── tts_webui.py        Web-Dienst
├── config.toml         Einstellungen
├── requirements.txt    Python-Pakete
├── drama-3/            Modell von Hugging Face (nicht im Repository)
├── voices/             Referenzclips (nicht im Repository)
└── data/               Aufträge und MP3s (nicht im Repository)
```

Liegt das Modell woanders, trägst du den Pfad in `config.toml` unter `model` ein.

## Bedienung

Auf der Seite ein Dokument wählen, der Auftrag startet sofort. Beim ersten Auftrag lädt das Sprachmodell einige Minuten, nach 10 Minuten Leerlauf wird es wieder beendet. Sind weniger als 85 GB VRAM frei, wartet der Auftrag, bis die GPU frei ist.

Fertige MP3s lassen sich direkt abspielen und herunterladen. Sie bleiben 30 Tage gespeichert.

## Stimmen

Fest eingebaute Stimmen gibt es lokal nicht, jede Stimme ist ein kurzer Referenzclip. Beim ersten Start lädt das Skript die Stimme `en0`, den Referenzclip aus dem Drama-3-README. Für eine weitere Stimme legst du zwei Dateien gleichen Namens in `voices/`:

```bash
ffmpeg -i aufnahme.m4a -ac 1 voices/anna.wav
echo "The exact words spoken in the clip." > voices/anna.txt
```

- Der Clip sollte in der Sprache des Textes sein, sonst dürfte sein Akzent mitkommen.
- 10 bis 30 Sekunden, ein Sprecher, keine Musik, höchstens 60 Sekunden.
- Die `.txt` enthält den exakten Wortlaut in UTF-8.
- Ab zwei Stimmen erscheint auf der Seite eine Auswahl, ohne Neustart.
- Fremde Stimmen nur mit Einwilligung der Person, das verlangt die Lizenz des Modells.

## Konfiguration

Alle Einstellungen stehen kommentiert in `config.toml`. Jeder Wert lässt sich per Umgebungsvariable überschreiben, der Name ist der Schlüssel in Großbuchstaben mit Präfix `TTS_`:

```bash
TTS_CONCURRENCY=64 TTS_PORT=9000 python tts_webui.py
```

Mit `concurrency` steuerst du die GPU-Last, die Seite zeigt dazu Auslastung und Tempo gegenüber Echtzeit. Bei weniger als 96 GB VRAM müssen `engine_mem` und `min_free_gib` angepasst werden.

## Als Dienst starten

`~/.config/systemd/user/audilekt.service`, Pfade anpassen:

```ini
[Service]
WorkingDirectory=/pfad/zu/audilekt
ExecStart=/pfad/zu/miniconda3/envs/audilekt/bin/python tts_webui.py
Restart=on-failure

[Install]
WantedBy=default.target
```

```bash
systemctl --user daemon-reload
systemctl --user enable --now audilekt
loginctl enable-linger "$USER"
```

## Fehlersuche

- **Modell startet nicht:** `tail -50 data/backend.log`, die Seite zeigt nur das Ende.
- **`pip check` meldet Konflikte** bei `protobuf`, `grpcio-*` und `decord`: Das ist erwartet und betrifft Drama 3 nicht.
- **`untested GPU sm_120` im Log:** Auf Blackwell normal, SGLang-Omni nutzt dort Rückfallpfade.

## Sicherheit

Es gibt keinen Login, und der Dienst lauscht auf allen Netzwerkschnittstellen. Betreibe ihn nur in einem vertrauenswürdigen Netz.

## Lizenzen

Der Code dieses Projekts steht unter der [MIT-Lizenz](LICENSE). Sie gilt nicht für das Modell und nicht für die installierten Python-Pakete.

Drama 3 steht unter der [Fish Audio Community License v1.0](https://huggingface.co/fishaudio/drama-3/blob/main/LICENSE). Maßgeblich ist der Lizenztext, kurz zusammengefasst:

- Forschung und nicht-kommerzielle Nutzung sind frei, kommerzielle Nutzung nur unter 500.000 US-Dollar Jahresumsatz.
- Die Lizenz deckt keinen „Hosted Service" ab. Audilekt Dritten als Sprachdienst zugänglich zu machen, braucht daher in der Regel eine eigene Vereinbarung mit Fish Audio.
- Die Stimme einer identifizierbaren Person darf nur mit deren Einwilligung nachgeahmt werden.
- Erzeugte Audiodaten dürfen nicht zum Training von Sprachgenerierungsmodellen verwendet werden.

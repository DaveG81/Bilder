#!/bin/bash

# Pfad zum Ordner mit dem Skript und der virtuellen Umgebung
SCRIPT_DIR="/Users/david/Bilder_Liste"

# Aktivieren des virtuellen Environments
source "$SCRIPT_DIR/myenv/bin/activate"

# Ausführen des Python-Skripts
python3 "$SCRIPT_DIR/Bilder_Liste.py"

# Deaktivieren des virtuellen Environments
deactivate


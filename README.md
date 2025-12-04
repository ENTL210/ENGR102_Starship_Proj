# 🚀 ENGR 102: Starship Problem Generator (Project Lux)

## 📖 Project Overview

Designed for the ENGR 102 curriculum at Oregon State University, this application is a Python-based data simulation engine that automates the creation of educational engineering datasets. It algorithmically generates high-fidelity, time-series operational data for autonomous delivery robots, allowing specific engineering trade-offs (e.g., mass vs. power consumption) to be modeled and analyzed by students.

## 💻 Tech Stack

- Python 3.12
- Poetry (dependency and environment management)
- `openpyxl` for Excel file generation
- JSON for configuration (`data.json`)

## 🧠 Key Technical Competencies

- **Stochastic Data Simulation**  
  Implements custom algorithms to generate randomized time-series data (07:30–22:00) that rigorously adheres to specific statistical constraints (Min, Max, and Target Average).

- **Automated File I/O**  
  Utilizes `openpyxl` to programmatically generate and format complex Excel (`.xlsx`) artifacts, creating 48 unique files per execution batch.

- **Configuration Management**  
  Uses JSON serialization (`data.json`) to decouple simulation parameters from the codebase, allowing for scalable difficulty adjustments.

- **Data Aggregation Logic**  
  Includes an automated **Answer Key** generator that parses generated Excel files, performs aggregate calculations, and exports a grading key, demonstrating valid Verification & Validation (V&V) logic.

---

## ⚙️ Environment Setup

This project uses **Poetry** for dependency management to ensure a reproducible build environment.

### Prerequisites

- Python **3.12.4** (ensure this is installed and added to your `PATH`)
- Poetry (if not installed, run: `pip install poetry`)

---

## 🪟 Windows Setup

Open **PowerShell** or **Command Prompt** as Administrator and navigate to the project folder.

```bash
# 1. Configure Poetry to use Python 3.12
poetry env use python3.12

# 2. Install dependencies defined in pyproject.toml
poetry install

# 3. Activate the virtual environment
poetry shell
```
## 🍎 MacOS / 🐧 Linux Setup

Open your **Terminal** and navigate to the project folder.

```bash
# 1. Configure Poetry to use Python 3.12
poetry env use 3.12

# 2. Install dependencies
poetry install

# 3. Activate the virtual environment
poetry shell
```
---

## 🧑‍💻 Running the Application

Once your shell is active (you should see the environment name in parentheses in your terminal), you can run the main program.

```bash
# 1. Start the CLI Tool
python main.py

# 2. Using the Menu:
---------------------------------------------------------------------------
Welcome to the ENGR102 Starship Problem Generator.
---------------------------------------------------------------------------
Options:
[1] Generate Problems
[2] Generate Answers Key
[3] Quit Tool
```

### Menu Options:
**Option 1**: Generate Problems — Prompts for an export directory and generate 4 random datasets (Winter, Spring, Summer, Fall) with 12 Excel files each.
<br/>
<br/>
**Option 2**: Generate Answers Key — Scans a directory containing the generated datasets and creates an ```engr102_starship_answers_key.txt``` file with calculated averages, sums, and trade-off percentages.

---

## 📂 Configuration
To adjust the simulation parameters (for example, changing the average temperature for Winter or the power consumption of the robot), modify the ```data.json``` file in the root directory. No code changes are required.

<br /> <p align="center"> Edward Lam — OSU ENGR 102 F2025 024 Final Project </p>


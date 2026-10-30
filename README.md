# Dog Racing – Excel VBA Betting Game 🐕🏁

🔗 **GitHub Repository:** https://github.com/Rogachichi/DogRacing-VBA.git

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Visual_Basic_for_Applications-blue?style=for-the-badge)

A fully functional, interactive Dog Racing betting game built entirely in Microsoft Excel using **Visual Basic for Applications (VBA)**. This project demonstrates logical programming, random number generation, dynamic data logging, and automated chart generation within the Excel environment.

## 🚀 Key Features

* **Interactive Betting System:** Users can set an initial budget, choose a dog to bet on (1-5), and place specific wager amounts via intuitive dialog boxes.
* **Race Simulation & Payouts:** A randomized algorithm determines the winning dog. The game handles logic for wins and losses, updating the balance automatically (features a 5.0x payout multiplier).
* **Automated Data Logging:** Generates a dedicated "Game History" (Historia gry) sheet to log round numbers, match outcomes, and remaining balance.
* **Dynamic Visualization:** Automatically plots a dynamic chart of the user's balance across all played rounds to visualize financial progression.
* **Session Management:** Built-in macros to restart the game, clear historical data, and reset the chart for a new session.

## 🛠️ Technology Stack

* **Environment:** Microsoft Excel
* **Language:** VBA (Visual Basic for Applications)
* **Mechanics:** Excel Object Model (Worksheets, Range, ChartObjects, MsgBox, InputBox)

## ⚙️ Setup & Installation

1. Clone this repository to your local machine.
2. Open the `DogRacing.xlsm` file in Microsoft Excel.
3. **Enable Macros:** Excel will display a security warning regarding macros. Click "Enable Content" to allow the VBA scripts to run.
4. Use the provided on-sheet buttons to start a new game, place bets, or reset the statistics.
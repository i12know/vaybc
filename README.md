# vaybc
VAYBC is a comprehensive Jeopardy-style game PowerPoint template and VBA code that's been customized for Bible Challenge competition at VAY Sports Fest

# VAY Bible Challenge 2025 - PowerPoint Game System

A customized Jeopardy-style PowerPoint game system specifically designed for VAY Sports Ministry's Bible Challenge competition. This interactive presentation allows for dynamic gameplay with real-time scoring, countdown timers, and automated question management.

**Two versions available:**
- **Practice:** Full 8-player, 6-category board for training
- **Official Game:** Competition-compliant 3-player, 5-category board

{TODO: Put photo here}
*Official game board showing 5 categories with Bible-themed questions*

## 🎯 Features

### Two Template Options

#### Official Game (VAY Competition Rules)
- **5 Bible Categories** with 5 difficulty levels each
- **3 Player Support** (VAY regulation compliance)
- Competition-ready interface optimized for official gameplay

#### Practice Game (For Training)
- **6 Categories** with 5 difficulty levels each  
- **8 Player Support** for larger practice groups
- Full original template functionality for comprehensive training

### Core Gameplay (Both Templates)
- **Daily Doubles** with random placement and custom wagering
- **Final Jeopardy** with individual player wagering
- **Interactive Game Board** with clickable question tiles
- **Real-time Score Tracking** with persistent scoreboards
- **Automatic Score Management** with audit trail logging

### Technical Features
- **Excel Import** - Load questions and answers from spreadsheet templates
- **Countdown Timer** - Visual timer for each question (customizable duration)
- **Score Adjustment Tools** - Manual score correction capabilities
- **Master Slide Integration** - Persistent scoreboards across all slides
- **Question Navigation** - Automatic slide transitions and board management

## 📋 Requirements

- **Microsoft PowerPoint** (2010 or later recommended)
- **Windows Operating System** (tested on Windows 10)
- **VBA/Macros Enabled** in PowerPoint
- **Microsoft Excel** (for question import functionality)

## 🚀 Quick Start

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/i12know/vaybc.git
   cd vaybc
   ```

2. **Choose your template:**
   - **For Official Competition:** `PowerPoint_Files/VAY_Bible_Challenge_Official.pptm` (3 players, 5 categories)
   - **For Practice/Training:** `PowerPoint_Files/VAY_Bible_Challenge_Practice.pptm` (8 players, 6 categories)

3. **Open the PowerPoint file:**
   - **Enable macros** when prompted by PowerPoint
   - Both templates use the same VBA codebase

4. **Configure player/team names:**
   - Go to **View > Slide Master**
   - **Official Template:** Edit team name text boxes (`TextBox 2`, `TextBox 3`, `TextBox 4`)
   - **Practice Template:** Edit team name text boxes (`TextBox 2` through `TextBox 9`)
   - Exit Slide Master view

### Basic Usage

1. **Setup Questions:**
   - Use the Excel template in `Templates/JeopardyQuestionTemplate.xlsx`
   - Fill in categories, values, questions, and answers
   - Run `ImportFromExcel()` macro to load questions

2. **Start Game:**
   - Press `F5` to start slideshow
   - Click question values on the board to navigate
   - Use player buttons to award/deduct points

3. **Game Controls:**
   - **Correct Answer:** Click green "Player X Correct" button
   - **Incorrect Answer:** Click red "Player X Incorrect" button
   - **Manual Score Adjustment:** Use `AdjustScores()` macro
   - **Reset Game:** Use `resetAll()` macro

## 🛠️ Customization

### Template Differences

The codebase supports both configurations through the same VBA constants:

```vba
Private Const PLAYER_COUNT = 8    # Full support for up to 8 players
Private Const CATEGORY_COUNT = 6  # Full support for up to 6 categories
```

**Official Template (VAY Competition Rules):**
- Uses only players 1-3 (scoreboards 4-8 hidden)
- Uses only categories 1-5 (category 6 hidden)
- Simplified interface for official competition

**Practice Template (Full Training):**
- All 8 player scoreboards visible and functional
- All 6 categories visible and functional  
- Complete original template experience

### Creating Custom Variations

**To modify existing templates:**
1. **Adjust visible elements** in Master Slide view
2. **Show/hide scoreboards** by modifying Master Slide layouts
3. **Add/remove player buttons** on question slides
4. **Update Excel templates** to match category count

### Template-Specific Functions

Both templates use the same VBA functions, but with different visual implementations:

```vba
' This function adapts based on visible players in each template
Private Function sNameOfPlayer(playerNo As Integer) As String
    ' Official template: Returns names for players 1-3 only
    ' Practice template: Returns names for players 1-8
End Function
```

### Adding Countdown Timer

The repository includes an enhanced timer system. To implement:

```vba
' Add to your question display routine
Sub StartCountdownTimer()
    ' Timer implementation - see VBA/TimerModule.bas
End Sub
```

### Excel Template Format

Questions should be formatted in Excel as:
- **Column A:** Category Name
- **Column B:** Point Value  
- **Column C:** Question Text
- **Column D:** Answer Text
- **Column E:** Additional Comments

**Template Files:**
- `JeopardyQuestionTemplate.xlsx` - Fill out 5 categories instead of 6 for official competitions

## 📁 Repository Structure

```
VAY-Bible-Challenge/
├── VBA/
│   ├── Module1.bas              # Main game logic (shared by both templates)
│   ├── TimerModule.bas          # Countdown timer implementation
│   └── UserForms/               # Score adjustment dialogs
├── PowerPoint_Files/
│   ├── VAY_Bible_Challenge_Official.pptm    # 3 players, 5 categories
│   └── VAY_Bible_Challenge_Practice.pptm    # 8 players, 6 categories
├── Templates/
│   └── JeopardyQuestionTemplate.xlsx    # Fill out 5 categories instead of 6 for official competitions
├── Documentation/
│   ├── PowerPoint_Structure.md   # Detailed slide layout for both versions
│   ├── Installation_Guide.md     # Step-by-step setup
│   └── Customization_Guide.md    # Advanced modifications
├── docs/
│   └── images/                   # Screenshots and diagrams
└── README.md
```

## 🎮 Game Rules & Scoring

### Template-Specific Rules

#### Official VAY Bible Challenge Rules
- **5 Categories:** Old Testament, New Testament, Characters, Geography, Prophecy
- **Point Values:** $100, $200, $300, $400, $500 per category, with double the value for Round 2
- **3 Teams Maximum:** Compliant with VAY competition regulations
- **Daily Doubles:** 1-2 per game, randomly placed in categories 1-5
- **Final Jeopardy:** All teams participate with wagering

#### Practice Template Rules  
- **6 Categories:** Includes an additional category for comprehensive training
- **Point Values:** $100, $200, $300, $400, $500 per category, with double the value for Round 2
- **8 Teams Maximum:** Full training environment
- **Daily Doubles:** 1-2 per game, randomly placed in categories 1-6
- **Final Jeopardy:** All participating teams wager

### Scoring System
- Correct answers **add** points
- Incorrect answers **subtract** points  
- Daily Double scoring based on the team's wager
- Final Jeopardy uses individual team wagers

## 🔧 Troubleshooting

### Common Issues

**Macros not working:**
- Ensure macros are enabled: File > Options > Trust Center > Macro Settings
- Check if VBA references are intact

**Excel import failing:**
- Verify Excel template format matches expected structure
- Check for special characters in questions/answers

### Debug Mode

Enable debug output by adding to VBA:
```vba
Debug.Print "Current player: " & player
Debug.Print "Score change: " & direction * iSlideValue
```

## 📝 Contributing

We welcome contributions! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Guidelines
- Document all VBA functions with comments
- Test changes on multiple PowerPoint versions
- Update documentation for new features
- Follow existing naming conventions

## 📜 License & Credits

### Original Work
**Created by Kevin Dufendach (2009-2016)**
- Original Jeopardy PowerPoint Template by krd.public+Jeopardy@gmail.com on http://sites.google.com/site/dufmedical (defunct)
**Updated by VAY (2022-present)**
- Rev.2022 by Bumble:
  - Reduced fonts, made it 5 columns (by moving the 6th off screen), changed title from “Jeopardy” to “Bible Challenge”
  - Added JeopardyTheme.mp3, Time-Up.wav to Slide Master Prompt & Final Jeopardy. JeopardyBoardFill.wav sound effect will play when you mouse-over the title screen (music files must be in the same folder)
  - Coded replacement of “~” in Excel import file to a new paragraph after “=======“ so that translation could be added more easily
  - Coded an Audit Trail of scoring on the last slide to support dispute resolution
  - Moved players 5-8 off-screen to make the 4 players version, then moved player 4 off-screen to make the 3 players version
- Rev.2023:
  - Hanh added “GLA SBC WCC” to the Master Slides 1&2 for Player 1, 2, 3 labels. Note: Replace the content of the text, do NOT delete the TextBox, and create a new TextBox
  - Bumble added the Final Jeopardy button in the Round 1 Board, Rules button, and distributed the music to each slide to make it more suitable for multimedia questions
  - Hanh modified Final BC Topic: Change Prompt to Topic, change Response to Final Prompt, add music, and Final Answer after the music stops
  - Bumble adjusted the Daily Doubles for 5 columns, added a check for duplicates
- Rev.2025:
  - Bumble created a GitHub for it.

### License
This project is licensed under the **GNU General Public License v3.0 or later**.

```
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
```

See [LICENSE](LICENSE) file for full terms.

### VAY Modifications
**Modified for VAY Sports Ministry Bible Challenge 2025**

**Official Template:**
- Reduced to 5 categories (compliance with VAY rules)
- Limited to 3 players maximum
- Streamlined interface for competition use

**Practice Template:**  
- Maintained full 6 categories and 8 players
- Complete training environment
- Preparation for competition scenarios

**Both Templates:**
- Customized for Bible-themed content
- Enhanced audit trail and scoring
- Shared VBA codebase for consistency

## 🤝 Acknowledgments

- **Kevin Dufendach** for the original Jeopardy PowerPoint template
- **VAY Sports Ministry** for sponsoring the Bible Challenge competition
- **Community contributors** who helped customize and test the system

## 📞 Support

For questions about this implementation:
- Open an issue in this repository
- Contact VAY Sports Ministry technical team

## 🎯 Future Enhancements

- [ ] Count-down timer with numbers displayed on screen
- [ ] "Revert" Button to revert the scoring should the judges approved a challenge

---

**Built with ❤️ for VAY Sports Ministry Bible Challenge**

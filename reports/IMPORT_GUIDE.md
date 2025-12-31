# Import Guide - I&C Project Management

## Import Overview

This guide explains how to import and set up the I&C Project Management solution in your environment.

**Last Updated**: 2025-12-30 20:37:13+08:00

---

## Prerequisites

### Required Software

- [x] **Microsoft Excel** (2016 or later)
  - For viewing and editing Excel templates
  - Microsoft 365 subscription recommended

- [x] **Python 3.7+** (Optional, for regenerating Excel files)
  - Download from: https://www.python.org/downloads/
  - Used for running automation scripts

- [x] **Microsoft 365 Account** (Optional, for full integration)
  - Teams
  - Planner
  - OneNote
  - Outlook
  - To Do

### Optional Software

- [ ] **Git** (for version control)
- [ ] **Visual Studio Code** (for editing Markdown files)
- [ ] **Power Automate** (for automation workflows)

---

## Quick Import (3 Minutes)

### Step 1: Download Project Files

```bash
# If using Git
git clone <repository-url>
cd "I&C Project Management"

# Or download ZIP and extract
```

### Step 2: Open Excel Files

```bash
# Open main dashboard
open Dashboard.xlsx

# Open task tracker
open Task_Tracker.xlsx

# Open weekly report template
open Weekly_Report_Template.xlsx
```

### Step 3: Start Using

1. Open `QUICK_START.md` for 3-minute quick start guide
2. Customize project information in Excel files
3. Begin managing your project!

---

## Detailed Import Process

### Phase 1: File System Setup

#### Step 1.1: Create Project Directory

```bash
# Create project directory
mkdir -p ~/Projects/IC-Project-Management
cd ~/Projects/IC-Project-Management
```

#### Step 1.2: Copy All Files

```bash
# Copy all project files
cp -r /path/to/download/* .

# Verify files
ls -la
```

**Expected Files**:
- 3 Excel files (*.xlsx)
- 4 Python scripts (*.py)
- 15+ Markdown files (*.md)
- 1 openspec directory

#### Step 1.3: Set Permissions

```bash
# Make Python scripts executable
chmod +x *.py

# Verify
ls -l *.py
```

### Phase 2: Excel Templates Import

#### Step 2.1: Verify Excel Files

```bash
# Check Excel files
ls -lh *.xlsx
```

**Expected Output**:
```
Dashboard.xlsx (15 KB)
Task_Tracker.xlsx (9 KB)
Weekly_Report_Template.xlsx (7 KB)
```

#### Step 2.2: Open and Customize

1. **Open Dashboard.xlsx**
   - Go to "專案總覽" sheet
   - Update project name (B1)
   - Update project manager (B2)
   - Update start date (B3)
   - Update end date (B4)

2. **Open Task_Tracker.xlsx**
   - Review all 4 sheets
   - Customize as needed

3. **Open Weekly_Report_Template.xlsx**
   - Review template structure
   - Customize report sections

#### Step 2.3: Upload to OneDrive (Optional)

1. Login to OneDrive: https://onedrive.live.com
2. Create folder "I&C Project Management"
3. Upload all Excel files
4. Set sharing permissions

### Phase 3: Python Scripts Setup (Optional)

#### Step 3.1: Install Python

```bash
# Check Python version
python3 --version

# Should be 3.7 or higher
```

#### Step 3.2: Install Dependencies

```bash
# Install openpyxl
pip3 install openpyxl

# Or use the auto-install script
python3 generate_all.py
```

#### Step 3.3: Test Scripts

```bash
# Test generation (will overwrite existing files)
python3 generate_all.py
```

**Expected Output**:
```
✅ 所有 Excel 檔案已成功建立！

📁 生成的檔案:
   • Dashboard.xlsx
   • Task_Tracker.xlsx
   • Weekly_Report_Template.xlsx
```

### Phase 4: Microsoft 365 Integration (Optional)

#### Step 4.1: Create Teams Team

1. Open Microsoft Teams
2. Click "Join or create a team"
3. Click "Create team"
4. Select "From scratch"
5. Choose "Private" or "Public"
6. Name: "I&C Project Management"
7. Add team members

#### Step 4.2: Create Channels

Create the following channels:
- 📋 一般 (General) - Default
- 🚀 專案啟動 (Project Initiation)
- 📊 任務追蹤 (Task Tracking)
- 📝 會議紀錄 (Meeting Notes)
- 📈 週報與檢討 (Weekly Review)
- ✅ 結案報告 (Project Closure)

#### Step 4.3: Add Excel Tabs

For each channel:
1. Click "+" to add a tab
2. Select "Excel"
3. Choose file from OneDrive
4. Click "Save"

#### Step 4.4: Create Planner Board

1. In Teams, click "+" to add a tab
2. Select "Planner"
3. Create new plan: "I&C Project Management - Master Plan"
4. Create buckets:
   - 🔵 待辦 (To Do)
   - 🟡 進行中 (In Progress)
   - 🟢 審核中 (Review)
   - 🟣 已完成 (Done)
   - 🔴 已封存 (Archived)

#### Step 4.5: Create OneNote Notebook

1. In Teams, click "+" to add a tab
2. Select "OneNote"
3. Create new notebook: "I&C Project Management Notebook"
4. Create sections:
   - 📌 專案章程
   - 🎯 專案目標與範圍
   - 👥 利害關係人
   - 📅 會議紀錄
   - 💡 設計文件
   - 📚 知識庫
   - ✅ 結案文件

### Phase 5: Power Automate Setup (Optional)

See `POWER_AUTOMATE_GUIDE.md` for detailed instructions.

**Workflows to Create**:
1. New Task Notification
2. Task Due Reminder
3. Weekly Report Auto-Generation
4. Task Completion Sync

---

## Import Checklist

### File System
- [ ] Project directory created
- [ ] All files copied
- [ ] Permissions set correctly

### Excel Templates
- [ ] Dashboard.xlsx opened and verified
- [ ] Task_Tracker.xlsx opened and verified
- [ ] Weekly_Report_Template.xlsx opened and verified
- [ ] Project information customized
- [ ] Sample data reviewed

### Python Scripts (Optional)
- [ ] Python 3.7+ installed
- [ ] openpyxl installed
- [ ] Scripts tested successfully

### Microsoft 365 (Optional)
- [ ] Teams team created
- [ ] Channels created
- [ ] Excel tabs added
- [ ] Planner board created
- [ ] OneNote notebook created

### Power Automate (Optional)
- [ ] Workflow 1: New Task Notification
- [ ] Workflow 2: Task Due Reminder
- [ ] Workflow 3: Weekly Report Auto-Generation
- [ ] Workflow 4: Task Completion Sync

---

## Troubleshooting

### Issue 1: Excel Files Won't Open

**Problem**: Excel files show error when opening

**Solutions**:
1. Ensure Microsoft Excel 2016 or later is installed
2. Enable macros if prompted
3. Check file permissions
4. Try opening in Excel Online

### Issue 2: Python Scripts Fail

**Problem**: `python3 generate_all.py` shows errors

**Solutions**:
```bash
# Check Python version
python3 --version

# Reinstall openpyxl
pip3 uninstall openpyxl
pip3 install openpyxl

# Run with verbose output
python3 -v generate_all.py
```

### Issue 3: Cannot Upload to OneDrive

**Problem**: Upload fails or is slow

**Solutions**:
1. Check internet connection
2. Verify OneDrive storage space
3. Try uploading files one by one
4. Use OneDrive desktop app

### Issue 4: Teams Integration Issues

**Problem**: Cannot add Excel tabs in Teams

**Solutions**:
1. Ensure files are in OneDrive
2. Check Teams permissions
3. Refresh Teams app
4. Try using Teams web version

### Issue 5: Power Automate Workflows Don't Trigger

**Problem**: Workflows created but not running

**Solutions**:
1. Check workflow is turned ON
2. Verify trigger conditions
3. Check connection permissions
4. Review run history for errors

---

## Import Validation

### Validation Checklist

Run these checks to ensure successful import:

#### Excel Files Validation

```bash
# Check files exist
ls -lh *.xlsx

# Open each file
open Dashboard.xlsx
open Task_Tracker.xlsx
open Weekly_Report_Template.xlsx
```

**Verify**:
- [ ] All 3 files open without errors
- [ ] All worksheets are present
- [ ] Formulas calculate correctly
- [ ] Charts display properly
- [ ] Conditional formatting works

#### Python Scripts Validation

```bash
# Test generation
python3 generate_all.py
```

**Verify**:
- [ ] No errors during execution
- [ ] All 3 Excel files generated
- [ ] Files have correct structure

#### Documentation Validation

```bash
# Check all documentation files
ls -lh *.md
```

**Verify**:
- [ ] All Markdown files present
- [ ] Files open correctly
- [ ] Links work properly

---

## Post-Import Configuration

### Customize for Your Project

#### 1. Update Project Information

In `Dashboard.xlsx` → "專案總覽":
- Project Name (B1)
- Project Manager (B2)
- Start Date (B3)
- End Date (B4)
- Project Status (B6)

#### 2. Update Team Members

In `Dashboard.xlsx` → "任務清單":
- Data Validation → Column C (負責人)
- Update dropdown list with your team members

#### 3. Update Sample Tasks

In `Dashboard.xlsx` → "任務清單":
- Delete sample tasks (rows 2-11)
- Add your actual tasks

#### 4. Customize Categories

In `Dashboard.xlsx` → "任務清單":
- Data Validation → Column I (類別)
- Update categories as needed

---

## Import Best Practices

### 1. Backup Before Customizing

```bash
# Create backup
mkdir backup
cp *.xlsx backup/
```

### 2. Version Control

```bash
# Initialize Git repository
git init
git add .
git commit -m "Initial import"
```

### 3. Document Changes

- Keep CHANGE_LOG.md updated
- Document customizations
- Track team feedback

### 4. Regular Backups

- Backup Excel files daily
- Backup to OneDrive
- Keep local copies

### 5. Team Training

- Share USER_GUIDE.md with team
- Conduct training session
- Create FAQ document

---

## Next Steps After Import

### Immediate (Today)

1. [ ] Review QUICK_START.md
2. [ ] Customize Excel templates
3. [ ] Share with team

### This Week

1. [ ] Complete Microsoft 365 setup
2. [ ] Create Planner tasks
3. [ ] Set up Power Automate workflows

### This Month

1. [ ] Train team members
2. [ ] Establish workflows
3. [ ] Gather feedback
4. [ ] Optimize processes

---

## Support Resources

### Documentation

- **QUICK_START.md** - 3-minute quick start
- **USER_GUIDE.md** - Complete user manual
- **EXECUTION_PLAN.md** - Detailed execution plan
- **POWER_AUTOMATE_GUIDE.md** - Automation guide

### Contact

- **Email**: benchen1981@smail.nchu.edu.tw
- **Project**: I&C Project Management

---

## Import Summary

**Total Time**: 15-30 minutes (basic import)  
**Total Time**: 4-5 hours (full setup with Microsoft 365)

**Files to Import**: 25+ files  
**Excel Templates**: 3 files  
**Python Scripts**: 4 files  
**Documentation**: 15+ files

---

**Guide Version**: 1.0  
**Last Updated**: 2025-12-30 20:37:13+08:00  
**Status**: ✅ Complete

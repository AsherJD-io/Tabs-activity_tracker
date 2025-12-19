@echo off
REM === Generic Auto Push Script ===
REM Use current folder as project directory

SET PROJECT_DIR=%cd%
SET REPO_NAME=%~n0
SET COMMIT_MSG=Auto-update

REM Initialize Git if not present
git rev-parse --is-inside-work-tree >nul 2>&1
IF ERRORLEVEL 1 (
    git init
    echo Initialized Git repository.
)

REM Add .gitignore if missing
IF NOT EXIST ".gitignore" (
    echo node_modules/ > .gitignore
    echo dist/ >> .gitignore
    echo .env >> .gitignore
)

git add .
git commit -m "%COMMIT_MSG%" 2>nul

REM Add remote if missing
git remote get-url origin >nul 2>&1
IF ERRORLEVEL 1 (
    git remote add origin git@github.com:AsherJD-io\%REPO_NAME%.git
)

git branch -M main
git push -u origin main

echo Done!
pause

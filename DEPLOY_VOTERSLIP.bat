@echo off
chcp 65001 >nul
title पंचायती राज चुनाव 2026 - सचित्र मतदाता पर्ची पोर्टल डिप्लॉय (voterslip/slip)
color 0B
cls

echo =========================================================================
echo   🗳️ सचित्र मतदाता पर्ची पोर्टल (With Pic) | GitHub डिप्लॉय इंजन
echo   Repository: https://github.com/voterslip/slip.git (main)
echo =========================================================================
echo.

cd /d "%~dp0"

echo [1/3] Git स्टेटस जाँचा जा रहा है...
git status -s

echo.
echo [2/3] कोड, डेटाबेस एवं सेटिंग्स कमिट की जा रही हैं...
git add -A
git commit -m "feat: complete voter slip portal with With-Pic tab, English A-Z sorting, 9/15/18/21 thermal layouts, and Google Sheet integration"

echo.
echo [3/3] GitHub (voterslip/slip) पर पुश जारी...
git push -u origin main --force

echo.
echo =========================================================================
echo   ✓ बधाई! सचित्र मतदाता पर्ची पोर्टल सफलतापूर्वक लाइव हो चुका है।
echo   लाइव यूआरएल: https://voterslip.github.io/slip/
echo =========================================================================
echo.
pause

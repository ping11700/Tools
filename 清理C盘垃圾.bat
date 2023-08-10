%1 mshta vbscript:CreateObject("Shell.Application").ShellExecute("cmd.exe","/c %~s0 ::","","runas",1)(window.close)&&exit

@echo off    
echo 正在清除系统垃圾文件，请稍等……    
del /f /s /q %systemdrive%\*.tmp    
del /f /s /q %systemdrive%\*._mp    
del /f /s /q %systemdrive%\*.log    
del /f /s /q %systemdrive%\*.gid    
del /f /s /q %systemdrive%\*.chk    
del /f /s /q %systemdrive%\*.old    
del /f /s /q %systemdrive%\recycled\*.*    
del /f /s /q %windir%\*.bak    
del /f /s /q %windir%\prefetch\*.*    
rd /s /q %windir%\temp & md %windir%\temp    
del /f /q %userprofile%\COOKIES s\*.*    
del /f /q %userprofile%\recent\*.*    
del /f /s /q "%userprofile%\Local Settings\Temporary Internet Files\*.*"    
del /f /s /q "%userprofile%\Local Settings\Temp\*.*"    
del /f /s /q "%userprofile%\recent\*.*"    
sfc /purgecache ‘清理系统盘无用文件    
defrag %systemdrive% -b ‘优化预读信息    
echo 清除系统LJ完成！    
echo. & pause 
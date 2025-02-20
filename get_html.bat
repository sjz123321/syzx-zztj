@echo off
setlocal EnableDelayedExpansion

:: 检查命令行参数
if "%~1"=="" (
    echo 错误：需要提供两个参数 ^(学期 周次^)
    echo 用法：
    echo   %~nx0 学期 周次
    echo   例如：%~nx0 1 4
    exit /b 1
)
if "%~2"=="" (
    echo 错误：需要提供两个参数 ^(学期 周次^)
    echo 用法：
    echo   %~nx0 学期 周次
    echo   例如：%~nx0 1 4
    exit /b 1
)

:: 设置参数值
set "new_xq=%~1"
set "new_zc=%~2"

:: 显示将使用的配置值
echo 使用的配置 - 学期：!new_xq!，周次：!new_zc!

.\curl-8.12.1_1-win64-mingw\bin\curl.exe "http://10.180.168.5/cnsyzx2009/zjgl/checksearch.asp" ^
  -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7" ^
  -H "Accept-Language: zh-CN,zh;q=0.9,en;q=0.8" ^
  -H "Cache-Control: max-age=0" ^
  -H "Connection: keep-alive" ^
  -H "Content-Type: application/x-www-form-urlencoded" ^
  -b "cookdate=; 101801685cnsyzx2009=LastPassword=ChXJBnM9971WF067&CookieDate=3&UserPassword=a9113aed05c5a059&UserName=%%CA%%A9%%BD%%E0%%D7%%EA; ASPSESSIONIDSSSBCQTA=OJLINHNCAKGBKOPKLEOOGFDF; username=%%CA%%A9%%BD%%E0%%D7%%EA; userpassword=" ^
  -H "Origin: http://10.180.168.5" ^
  -H "Referer: http://10.180.168.5/cnsyzx2009/zjgl/checksearch.asp" ^
  -H "Upgrade-Insecure-Requests: 1" ^
  -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36" ^
  --data-raw "xn=2024&xq=!new_xq!&zc=!new_zc!&njbj=%%B8%%DF%%B6%%FE%%C4%%EA%%BC%%B614%%B0%%E0&lou=&Submit=%%B2%%E9%%D1%%AF" ^
  --insecure > response!new_xq!-!new_zc!.html

echo 请求完成，学期：!new_xq!，周次：!new_zc!

endlocal
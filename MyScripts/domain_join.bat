@echo off
wmic computersystem where name="%computername%" call joindomainorworkgroup fjoinoptions=3 name="epsoftinc.global" username="EPSOFTINC\Administrator" Password="PK$RD0bs@"

shutdown /r /t 30
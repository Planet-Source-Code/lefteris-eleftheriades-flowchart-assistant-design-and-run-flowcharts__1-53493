@echo Off
@Echo Compressing at maximum with upx everything but the icons
@Echo You need UPX 1.24 to pack the Compiled Executeable
@Echo Give it a minute to look for the best compression...
upx --best --crp-ms=999999 --nrv2b --overlay=strip --compress-icons=0 FlowChart3.exe
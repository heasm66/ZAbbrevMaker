# ZAbbrevMaker
Tool to make better abbreviations for ZIL and the z-machine. Input is the zap-files that results from compiling the zil-files. Typical workflow:

    zilf.exe game.zil
    del game_freq.zap
    ZAbbrevMaker.exe > game_freq.xzap
    zapf.exe game.zap
    
Precompiled binaries for linux-arm, linux-x64, osx-x64, win-x86, win-x64 & win-arm in the folder /ZAbbrevMaker/bin/SingleFile/.

## Instructions

    ZAbbrevMaker 0.2
    ZAbbrevMaker [switches] [path-to-game]
    
     -l nn        Maxlength of abbreviations (default = 20)
     -f           Fast rounding. Try variants (add remove space) to abbreviations
                  and see if it gives better savings on account of z-chars rounding.
     -d           Deep rounding. Try up yo 10,000 variants from discarded abbreviations
                  and see if it gives better savings on account of z-chars rounding.
     -df          Try deep rounding and then fast rounding, in that order (default).
     -fd          Try fast rounding and then deep rounding, in that order.
     -r3          Always round to 3 for fast and deep rounding. Normally rounding
                  to 6 is used for strings stored in high memory for z4+ games.
     -b           Throw all abbreviations that have lower score than last pick back on heap.
                  (This only occasionally improves the result, use sparingly.)
     -v           Verbose. Prints progress and intermediate results.
     path-to-game Use this path. If omitted the current path is used.
    
    ZAbbrevMaker executed without any switches in folder with zap-files is
    the same as 'ZAbbrevMaker -l 20 -df'.
    
## References
https://intfiction.org/t/highly-optimized-abbreviations-computed-efficiently/48753  
https://gitlab.com/russotto/zilabbrs  
https://github.com/hlabrand/retro-scripts

REM Lower volume of mp3
D:\ffmpeg\bin\ffmpeg -y -i Mr_Sunny_Face.mp3 -filter:a "volume=0.1" R:\temp\low_music.mp3
REM Concat twice mp3
D:\ffmpeg\bin\ffmpeg -y -i "concat:R:\temp\low_music.mp3|R:\temp\low_music.mp3|R:\temp\low_music.mp3" -acodec copy R:\temp\low_music_x2.mp3

REM convert mp4 to mkv
D:\ffmpeg\bin\ffmpeg -y -i Tutorial_01\result\video.mp4 -codec copy R:\temp\input.mkv
REM mix mp3 with already exixting audio
D:\ffmpeg\bin\ffmpeg -y -i R:\temp\input.mkv -i R:\temp\low_music_x2.mp3 -filter_complex "[0:a][1:a]amerge=inputs=2[a]" -map 0:v -map "[a]" -c:v copy -c:a libvorbis -ac 2 -shortest R:\temp\output.mkv

REM Add audio track
REM D:\ffmpeg\bin\ffmpeg -i R:\temp\video.mkv -i D:\Music\Mr_Sunny_Face.mp3 -map 0:v -map 0:a -map 1:a -c copy R:\temp\output.mkv

pause
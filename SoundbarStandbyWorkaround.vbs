Set Sound = CreateObject("WMPlayer.OCX")
Sound.URL = "C:\Users\Peter\Music\WakeupHum3sAlternative.wav"
Sound.Controls.play
do while Sound.currentmedia.duration = 0
wscript.sleep 100
loop
wscript.sleep (int(Sound.currentmedia.duration)+1)*1000
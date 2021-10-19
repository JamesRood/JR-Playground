#!ps
#Speech
#timeout=300000
#https://www.scriptinglibrary.com/languages/powershell/powershell-text-to-speech/
#choose / install different voices
# - have not successfully gotten to work yet
#https://support.office.com/en-us/article/how-to-download-text-to-speech-languages-for-windows-10-d5a6b612-b3ae-423f-afa5-4f6caf1ec5d3
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {
  // f(), g(), ... are unused COM method slots. Define these if you care
  int f(); int g(); int h(); int i();
  int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
  int j();
  int GetMasterVolumeLevelScalar(out float pfLevel);
  int k(); int l(); int m(); int n();
  int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
  int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
  int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
  int f(); // Unused
  int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
public class Audio {
  static IAudioEndpointVolume Vol() {
    var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
    IMMDevice dev = null;
    Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
    IAudioEndpointVolume epv = null;
    var epvid = typeof(IAudioEndpointVolume).GUID;
    Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
    return epv;
  }
  public static float Volume {
    get {float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v;}
    set {Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty));}
  }
  public static bool Mute {
    get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
    set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
  }
}
'@
function Set-Volume{
    param (
        [Parameter(Mandatory = $true)]
        [float] $volume,
        [Parameter(Mandatory = $true)]
        [bool] $muted
    )
    [Audio]::Mute = $muted;
    [Audio]::Volume = $volume;
}
function Basic-Speech{
    param (
        [Parameter(Mandatory = $true)]
        [string] $phrase
    )
    Add-Type -AssemblyName System.Speech;
    $speechSynth = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer;
    $speechSynth.SelectVoice("Microsoft Zira Desktop");
    $speechSynth.Speak($phrase);
}
function Medium-Speech{
    param (
        [Parameter(Mandatory = $true)]
        [string] $phrase,
        [int] $rate = 2,
        [string] $voice
    )
    Add-Type -AssemblyName System.Speech;
    $robot = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer;
    #$robot.GetInstalledVoices() | foreach{$_.VoiceInfo};
    #Check input to see if it matches an installed voice pack.
    #If not, choose my default.
    $defaultvoice = $false;
    if($voice -eq '' -or $voice -eq $null){
        $defaultvoice = $true;
    }
    else{
        $voices = $robot.GetInstalledVoices();
        $voiceNotFound = $true;
        foreach($v in $voices){
            if($v.voiceinfo.name -eq $voice){
                $robot.SelectVoice($voice);
                $voiceNotFound = $false;
            }
        }
        if($voiceNotFound -eq $true){
            'Voice "{0}" not found' -f $voice;
            $defaultvoice = $true;
        }
    }
    if($defaultvoice){
        #'Default Voice chosen';
        $robot.SelectVoice("Microsoft Zira Desktop");
    }
    
    #clamp the rate of speech to standards
    if($rate -gt 10){
        $rate = 10;
    }
    elseif($rate -lt -10){
        $rate = -10;
    }
    $robot.Rate = $rate;
    #Speak line
    $robot.Speak($phrase);
}
#Cache Audio settings before script
$tempVol = [Audio]::Volume;
$tempMute = [Audio]::Mute;
#Set-Volume -volume 0.9 -muted $false;
#Basic-Speech -phrase 'USB Drive one two three four is too small.';
#Medium-Speech -phrase 'USB Drive one two three four is too small.' -rate 2 -voice 'Microsoft David Desktop';
$timeOfDay = 'Day';
switch((Get-Date).hour){
    {1..11 -contains $_}{
        $timeOfDay = 'Morning';
        break;
    }
    {12..17 -contains $_}{
        $timeOfDay = 'Afternoon';
        break;
    }
    {18..23 -contains $_}{
        $timeOfDay = 'Evening';
        break;
    }
    default{
        $timeOfDay = 'Day';
    }
}

$whoami = whoami
$whoami_formatted = $whoami -replace ".*\\",""
$User = $whoami_formatted -replace "\."," "

#Download Song
#$Download = New-Object System.Net.WebClient;
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Path = "C:\WINDOWS\Temp\rickedit2.wav";
#this line doesnt directly work because it needs to play a wav not an mp3
#I edited it with audacity and put it on our LT share for downloading
$songPath = "https://resources.complete-it.co.uk/Uploads/SuperSecretFolder/supersecretgoodbyesong.wav";
#$Download.DownloadFile($songPath, $Path);
Invoke-WebRequest -Uri $songPath -Outfile $path
Set-Volume -volume 0.5 -muted $false;
$PlayWav=New-Object System.Media.SoundPlayer;
$PlayWav.SoundLocation=$Path;
$companyName = "Jar Roo Computing";
#phone Number needs to be a strong with spaces between the numbers
#911 will be read as nine hundred eleven, vs 9 1 1 will be read correctly
$phoneNumber = "Your IT Support Company";
$message = "Attention Please...Attention Please...";
Medium-Speech -phrase $message -rate 2;
$message = "Good $timeOfDay. What can I say? My time has come. Who knows if things like this will continue once I have faded from existence. ";
Medium-Speech -phrase $message -rate 2;
$message = "If you are receiving this message, you have been deemed worthy of receiving this wonderful piece of music and someone who it has been an absolute pleasure working with.";
Medium-Speech -phrase $message -rate 2;
$message = "With that in mind, I do hope I have not caught you at an in-opportune time. If I have, please kill powershell.exe to stop the music.";
Medium-Speech -phrase $message -rate 2;
$message = "However, I hope you enjoy listening to it as much as I enjoyed making it. It's a bit of a long boy, but hopefully you'll agree it captures my spirit.";
Medium-Speech -phrase $message -rate 2;
$message = "Goodnight and god bless. I will miss you.";
Medium-Speech -phrase $message -rate 2;
#music
Set-Volume -volume 0.5 -muted $false;
$PlayWav.playsync();
#Reset Audio settings after script
Set-Volume -volume $tempVol -muted $tempMute;

import getopt, sys, win32com.client


options = "lo:v:r:i:ps:"
def usage():
    print("[usage] %s -l | [-o] [-v] [-r] (-i | -p | -s | text)" % sys.argv[0])
    print("    -l language: case insensitive, begins-with match")
    print("    -o output.wav")
    print("    -v voice: case insensitive, 'Microsoft' can be dropped.")
    print("    -r rate: -10 (slow) ... 10 (fast)")
    print("    -i input.txt")
    print("    -p sym: SAPI TTS XML <pron>")
    print("    -s sapi|ups|ipa ph: SSML <phoneme> (requires -v)")
    exit(1)

_sapi = win32com.client.Dispatch("SAPI.SpVoice")
_cat  = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
_cat.SetID(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", False)

def speak(voice, text):
    old = _sapi.Voice
    if voice: _sapi.Voice = voice
    try:
        _sapi.Speak(text)
    finally:
        if voice: _sapi.Voice = old

def saveas(wav, f):
    fs = win32com.client.Dispatch("SAPI.SpFileStream")
    fs.Open(wav, 3)
    old = _sapi.AudioOutputStream
    _sapi.AudioOutputStream = fs
    try:
        f()
    finally:
        fs.Close()
        _sapi.AudioOutputStream = old

def save(voice, text, wav):
    saveas(wav, lambda: speak(voice, text))

def getvoices():
    return _cat.EnumerateTokens()

def getvoice(name, quit=False):
    if name: name = name.lower()
    def check(t):
        n = t.GetAttribute("Name").lower()
        return n == name or n == "microsoft " + name
    voices = [t for t in getvoices() if check(t)]
    if voices: return voices[0]
    if quit:
        print("voice not found:", name)
        exit(1)
    return None

def showvoices(voices, quit=False):
    langs = [l.lower() for l in voices]
    def f(v):
        c = getlocale(v)
        n = v.GetAttribute("Name")
        d = v.GetDescription().split(" - ")
        return (c, n) if len(d) < 2 else (c + ", " + d[1], n)
    voices = [
        (l, d)
        for l, d in map(f, getvoices())
        if not langs or [la for la in langs if l.lower().startswith(la)]]
    for l, n in sorted(voices): print(l + ":", n)
    if quit: exit(0)

def setrate(rate, quit=False):
    if rate < -10 or rate > 10:
        print("rate is out of range: %d" % rate)
        if quit: exit(1)
    else:
        _sapi.Rate = rate

def getlocale(voice):
    ret = voice.id.split("\\")[-1].split("_")[2]
    return ret if ret[2] == "-" else ret[:2] + "-" + ret[2:]

def pron(*texts, sep=""):
    return "".join(['<pron sym="%s"/>%s' % (text, sep) for text in texts])

def ssml(lang, alph, *texts, sep=""):
    ret = '<speak version="1.0" xml:lang="%s">\n' % lang
    for text in texts:
        ret += '<phoneme alphabet="%s" ph="%s"/>%s\n' % (alph, text, sep)
    ret += '</speak>'
    return ret

if __name__ == "__main__":
    voice  = None
    output = None
    mkxml  = lambda texts: " ".join(texts)
    text   = None
    alph   = None
    prefix = suffix = ""
    try:
        opts, args = getopt.getopt(sys.argv[1:], options)
    except getopt.GetoptError as e:
        print(e)
        usage()
    for opt, optarg in opts:
        if   opt == "-l": showvoices(args, quit=True)
        elif opt == "-o": output = optarg
        elif opt == "-v": voice = getvoice(optarg, quit=True)
        elif opt == "-r": setrate(int(optarg), quit=True)
        elif opt == "-i":
            with open(optarg, encoding="utf-8") as f:
                text = f.read()
        elif opt == "-p": mkxml = lambda texts: pron(*texts)
        elif opt == "-s":
            if not optarg in ["sapi", "ups", "ipa"]:
                print("option -s is invalid")
                usage()
            alph = optarg
    if alph:
        if not voice:
            print("option -s requires -v")
            usage()
        mkxml = lambda texts: ssml(getlocale(voice), alph, *texts)
    if not text: text = mkxml(args)
    if not text: usage()

    if output:
        save(voice, text, output)
    else:
        speak(voice, text)
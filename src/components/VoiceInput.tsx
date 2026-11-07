import React, { useState, useEffect, useRef } from 'react';
import { Mic, MicOff, Sparkles, Globe, CornerDownLeft, Volume2, RotateCcw } from 'lucide-react';

interface VoiceInputProps {
  onExecute: (instruction: string) => void;
  isLoading: boolean;
  activeTableName?: string;
  onUndo?: () => void;
  canUndo?: boolean;
}

export const VoiceInput: React.FC<VoiceInputProps> = ({
  onExecute,
  isLoading,
  activeTableName,
  onUndo,
  canUndo
}) => {
  const [isListening, setIsListening] = useState(false);
  const [transcript, setTranscript] = useState('');
  const [language, setLanguage] = useState<'hi-IN' | 'en-IN' | 'en-US'>('hi-IN');
  const [speechError, setSpeechError] = useState<string | null>(null);

  const recognitionRef = useRef<any>(null);

  useEffect(() => {
    // Check Web Speech API support
    const SpeechRecognition =
      (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;

    if (SpeechRecognition) {
      const recognition = new SpeechRecognition();
      recognition.continuous = true;
      recognition.interimResults = true;
      recognition.lang = language;

      recognition.onresult = (event: any) => {
        let currentTranscript = '';
        for (let i = event.resultIndex; i < event.results.length; i++) {
          currentTranscript += event.results[i][0].transcript;
        }
        setTranscript(prev => (prev ? `${prev} ${currentTranscript.trim()}` : currentTranscript.trim()));
      };

      recognition.onerror = (event: any) => {
        console.error('Speech recognition error:', event.error);
        if (event.error === 'not-allowed') {
          setSpeechError('Microphone permission denied. Please allow microphone access in your browser.');
        } else {
          setSpeechError(`Speech error: ${event.error}`);
        }
        setIsListening(false);
      };

      recognition.onend = () => {
        setIsListening(false);
      };

      recognitionRef.current = recognition;
    } else {
      setSpeechError('Web Speech API is not supported in this browser. Please type commands directly.');
    }

    return () => {
      if (recognitionRef.current) {
        try {
          recognitionRef.current.stop();
        } catch (_) {}
      }
    };
  }, [language]);

  const toggleListening = () => {
    if (isListening) {
      try {
        recognitionRef.current?.stop();
      } catch (_) {}
      setIsListening(false);
    } else {
      setSpeechError(null);
      try {
        recognitionRef.current?.start();
        setIsListening(true);
      } catch (err: any) {
        console.error('Failed to start speech recognition:', err);
        setSpeechError('Failed to start speech recognition. Please type your instruction.');
        setIsListening(false);
      }
    }
  };

  const handleQuickCommand = (cmd: string) => {
    setTranscript(cmd);
  };

  const handleFormSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!transcript.trim()) return;
    if (isListening) {
      try {
        recognitionRef.current?.stop();
      } catch (_) {}
      setIsListening(false);
    }
    onExecute(transcript);
  };

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-lg">
      <div className="flex flex-wrap items-center justify-between gap-4 mb-3">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
            <Volume2 className="w-4 h-4" />
          </div>
          <div>
            <h2 className="text-sm font-semibold text-slate-100">Voice & Text Controller</h2>
            <p className="text-xs text-slate-400">Speak or type commands in Hindi, English, or Hinglish</p>
          </div>
        </div>

        {/* Language Selector & Undo Action */}
        <div className="flex items-center gap-2">
          <div className="flex items-center gap-1.5 px-2.5 py-1 bg-slate-800/80 border border-slate-700/60 rounded-lg text-xs text-slate-300">
            <Globe className="w-3.5 h-3.5 text-slate-400" />
            <select
              value={language}
              onChange={e => setLanguage(e.target.value as any)}
              className="bg-transparent text-slate-200 text-xs focus:outline-none cursor-pointer"
            >
              <option value="hi-IN" className="bg-slate-900 text-slate-100">हिन्दी / Hinglish (hi-IN)</option>
              <option value="en-IN" className="bg-slate-900 text-slate-100">English (India - en-IN)</option>
              <option value="en-US" className="bg-slate-900 text-slate-100">English (US - en-US)</option>
            </select>
          </div>

          {canUndo && onUndo && (
            <button
              onClick={onUndo}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded-lg text-xs font-medium transition-colors cursor-pointer"
              title="Undo last action"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Undo Step</span>
            </button>
          )}
        </div>
      </div>

      {speechError && (
        <div className="mb-3 px-3 py-2 bg-amber-950/40 border border-amber-800/60 rounded-lg text-xs text-amber-200">
          {speechError}
        </div>
      )}

      {/* Main Input Form */}
      <form onSubmit={handleFormSubmit} className="space-y-3">
        <div className="relative">
          <textarea
            value={transcript}
            onChange={e => setTranscript(e.target.value)}
            placeholder="Type or speak, e.g. 'Orders aur Prices me Item se VLOOKUP lagao aur Unit_Price le aao'..."
            rows={2}
            className="w-full pl-4 pr-24 py-3 bg-slate-950/70 border border-slate-800 rounded-lg text-sm text-slate-100 placeholder-slate-500 focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 resize-none font-sans"
          />

          <div className="absolute right-2.5 top-2.5 flex items-center gap-1.5">
            {/* Mic Button */}
            <button
              type="button"
              onClick={toggleListening}
              className={`p-2.5 rounded-lg font-medium transition-all cursor-pointer flex items-center justify-center ${
                isListening
                  ? 'bg-rose-500 text-white animate-pulse shadow-lg shadow-rose-500/20'
                  : 'bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700'
              }`}
              title={isListening ? 'Click to stop listening' : 'Click to start voice recognition'}
            >
              {isListening ? <MicOff className="w-4 h-4" /> : <Mic className="w-4 h-4" />}
            </button>

            {/* Run Button */}
            <button
              type="submit"
              disabled={isLoading || !transcript.trim()}
              className="px-3.5 py-2.5 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-800 disabled:text-slate-500 text-white rounded-lg text-xs font-medium transition-colors flex items-center gap-1 cursor-pointer disabled:cursor-not-allowed shadow-sm"
            >
              {isLoading ? (
                <div className="w-4 h-4 border-2 border-white/20 border-t-white rounded-full animate-spin" />
              ) : (
                <>
                  <span>Run</span>
                  <CornerDownLeft className="w-3.5 h-3.5" />
                </>
              )}
            </button>
          </div>
        </div>

        {/* Live Audio Status Indicator when listening */}
        {isListening && (
          <div className="flex items-center gap-3 px-3 py-2 bg-emerald-950/30 border border-emerald-900/50 rounded-lg text-xs text-emerald-300">
            <span className="flex h-2 w-2 relative">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
            </span>
            <span className="font-medium">Listening ({language})... Speak your instruction now</span>
            <span className="text-slate-400 text-xs ml-auto">Click red mic to finish</span>
          </div>
        )}

        {/* Quick Suggestion Chips */}
        <div className="pt-1">
          <div className="flex items-center gap-2 mb-2 text-xs text-slate-400">
            <Sparkles className="w-3.5 h-3.5 text-amber-400" />
            <span>Try these sample Hindi & English commands:</span>
          </div>
          <div className="flex flex-wrap gap-2">
            {[
              {
                label: '🔗 VLOOKUP Unit_Price',
                cmd: 'Orders aur Prices me Item aur Item_Name se VLOOKUP lagao aur Unit_Price le aao'
              },
              {
                label: '📊 Sort Quantity Descending',
                cmd: 'Orders ko Quantity ke hisab se bade se chota sort karo'
              },
              {
                label: '🤝 Common Products in Inventory',
                cmd: 'Orders aur Inventory me common product nikalo'
              },
              {
                label: '➕ Total = Quantity * Unit_Price',
                cmd: 'Add calculated column Total = Quantity * Unit_Price'
              },
              {
                label: '🏙️ Pivot Table by City',
                cmd: 'City ke hisab se Quantity ka total sum pivot banao'
              },
              {
                label: '🧹 Remove Duplicates',
                cmd: 'Duplicate rows hata do'
              }
            ].map((item, idx) => (
              <button
                key={idx}
                type="button"
                onClick={() => handleQuickCommand(item.cmd)}
                className="px-2.5 py-1 bg-slate-800/90 hover:bg-slate-700/90 text-slate-300 hover:text-white border border-slate-700/60 rounded-md text-xs transition-colors cursor-pointer text-left whitespace-nowrap"
              >
                {item.label}
              </button>
            ))}
          </div>
        </div>
      </form>
    </div>
  );
};

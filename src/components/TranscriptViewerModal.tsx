import React, { useState, useMemo } from 'react';
import {
  FileText,
  Clock,
  Users,
  Search,
  Copy,
  Check,
  Download,
  Filter,
  X,
  Sparkles,
  ExternalLink,
  Code,
  ShieldCheck,
  UserCheck
} from 'lucide-react';
import { TranscriptContent, TranscriptMetadata, TranscriptCue, IdentifiedSpeaker } from '../types/transcript';

interface TranscriptViewerModalProps {
  isOpen: boolean;
  onClose: () => void;
  meetingTitle: string;
  metadata?: TranscriptMetadata;
  content?: TranscriptContent;
  onlineMeetingId?: string;
}

export const TranscriptViewerModal: React.FC<TranscriptViewerModalProps> = ({
  isOpen,
  onClose,
  meetingTitle,
  metadata,
  content,
  onlineMeetingId,
}) => {
  const [selectedSpeaker, setSelectedSpeaker] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [activeTab, setActiveTab] = useState<'cues' | 'raw'>('cues');
  const [copied, setCopied] = useState<boolean>(false);

  // Formatage de durée (secondes -> mm:ss ou hh:mm:ss)
  const formatDuration = (totalSeconds: number): string => {
    const s = Math.round(totalSeconds);
    const hrs = Math.floor(s / 3600);
    const mins = Math.floor((s % 3600) / 60);
    const secs = s % 60;
    if (hrs > 0) {
      return `${hrs}h ${mins.toString().padStart(2, '0')}m ${secs.toString().padStart(2, '0')}s`;
    }
    return `${mins}m ${secs.toString().padStart(2, '0')}s`;
  };

  // Filtrage combiné : par intervenant et par requête de recherche
  const filteredCues = useMemo(() => {
    if (!content?.cues) return [];
    let list = content.cues;

    if (selectedSpeaker) {
      list = list.filter((c) => c.speakerName.toLowerCase() === selectedSpeaker.toLowerCase());
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      list = list.filter((c) => c.text.toLowerCase().includes(q) || c.speakerName.toLowerCase().includes(q));
    }

    return list;
  }, [content, selectedSpeaker, searchQuery]);

  const handleCopyRaw = async () => {
    if (!content?.rawContent) return;
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(content.rawContent);
      } else {
        const textarea = document.createElement('textarea');
        textarea.value = content.rawContent;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 3000);
    } catch (e) {
      console.warn('[Copy warning]', e);
    }
  };

  const handleDownloadVtt = () => {
    if (!content?.rawContent) return;
    const blob = new Blob([content.rawContent], { type: 'text/vtt;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const sanitizedTitle = meetingTitle.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 40);
    a.download = `Transcription_Teams_${sanitizedTitle}.vtt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-md flex items-center justify-center p-3 sm:p-5">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-4xl w-full flex flex-col max-h-[92vh] shadow-2xl overflow-hidden animate-in fade-in duration-200">
        {/* 1. En-tête Modal */}
        <div className="p-4 sm:p-5 border-b border-slate-800 flex items-start justify-between gap-3 bg-slate-950/60">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2 py-0.5 text-[10px] font-semibold bg-emerald-950 text-emerald-300 border border-emerald-800 rounded-md flex items-center gap-1 shadow-sm">
                <ShieldCheck className="w-3 h-3 text-emerald-400" />
                Microsoft Teams Graph v1.0
              </span>
              {metadata?.id && (
                <span className="text-[10px] font-mono text-slate-400 bg-slate-900 px-2 py-0.5 rounded border border-slate-800">
                  ID : {metadata.id.slice(0, 16)}...
                </span>
              )}
            </div>
            <h3 className="text-base font-bold text-white tracking-tight leading-snug">
              {meetingTitle}
            </h3>
            <p className="text-xs text-slate-400">
              Transcription capturée le{' '}
              {metadata?.createdDateTime
                ? new Date(metadata.createdDateTime).toLocaleDateString('fr-FR', {
                    day: 'numeric',
                    month: 'long',
                    year: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit',
                  })
                : 'Date non précisée'}
            </p>
          </div>

          <button
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg transition-colors shrink-0"
            title="Fermer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* 2. Bandeau métriques de transcription */}
        {content && (
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 p-3 sm:p-4 bg-slate-950/40 border-b border-slate-800/80 text-xs">
            <div className="p-2.5 rounded-xl bg-slate-900/60 border border-slate-800/70">
              <span className="text-slate-500 block text-[11px] flex items-center gap-1.5">
                <Clock className="w-3.5 h-3.5 text-blue-400" />
                Durée totale
              </span>
              <span className="text-white font-bold text-sm">
                {formatDuration(content.durationSeconds)}
              </span>
            </div>

            <div className="p-2.5 rounded-xl bg-slate-900/60 border border-slate-800/70">
              <span className="text-slate-500 block text-[11px] flex items-center gap-1.5">
                <FileText className="w-3.5 h-3.5 text-emerald-400" />
                Segments / Répliques
              </span>
              <span className="text-white font-bold text-sm">
                {content.cues.length}
              </span>
            </div>

            <div className="p-2.5 rounded-xl bg-slate-900/60 border border-slate-800/70">
              <span className="text-slate-500 block text-[11px] flex items-center gap-1.5">
                <Users className="w-3.5 h-3.5 text-purple-400" />
                Locuteurs identifiés
              </span>
              <span className="text-white font-bold text-sm">
                {content.speakersCount}
              </span>
            </div>

            <div className="p-2.5 rounded-xl bg-slate-900/60 border border-slate-800/70">
              <span className="text-slate-500 block text-[11px] flex items-center gap-1.5">
                <Sparkles className="w-3.5 h-3.5 text-amber-400" />
                Volume de mots
              </span>
              <span className="text-white font-bold text-sm">
                {content.totalWords.toLocaleString('fr-FR')}
              </span>
            </div>
          </div>
        )}

        {/* 3. Bandeau des Locuteurs Identifiés & Barre de recherche */}
        <div className="p-3 sm:p-4 border-b border-slate-800/80 bg-slate-900/60 space-y-3">
          {/* Locuteurs */}
          {content?.speakers && content.speakers.length > 0 && (
            <div className="space-y-1.5">
              <div className="flex items-center justify-between text-xs">
                <span className="text-slate-400 font-semibold flex items-center gap-1.5">
                  <UserCheck className="w-3.5 h-3.5 text-indigo-400" />
                  Filtrer par locuteur convoqué :
                </span>
                {selectedSpeaker && (
                  <button
                    onClick={() => setSelectedSpeaker(null)}
                    className="text-[11px] text-blue-400 hover:text-blue-300 font-medium underline"
                  >
                    Réinitialiser le filtre
                  </button>
                )}
              </div>
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  onClick={() => setSelectedSpeaker(null)}
                  className={`px-2.5 py-1 text-xs rounded-lg border font-medium transition-all ${
                    selectedSpeaker === null
                      ? 'bg-blue-600 text-white border-blue-500 shadow-sm'
                      : 'bg-slate-950 text-slate-300 border-slate-800 hover:border-slate-700'
                  }`}
                >
                  Tous ({content.cues.length})
                </button>
                {content.speakers.map((spk) => {
                  const isSelected = selectedSpeaker === spk.name;
                  return (
                    <button
                      key={spk.name}
                      onClick={() => setSelectedSpeaker(isSelected ? null : spk.name)}
                      className={`px-2.5 py-1 text-xs rounded-lg border font-medium flex items-center gap-1.5 transition-all ${
                        isSelected
                          ? 'bg-indigo-600 text-white border-indigo-500 shadow-sm'
                          : 'bg-slate-950 text-slate-300 border-slate-800 hover:border-slate-700'
                      }`}
                    >
                      <span className={`w-2 h-2 rounded-full bg-gradient-to-r ${spk.avatarColor || 'from-blue-500 to-indigo-700'}`} />
                      <span>{spk.name}</span>
                      <span className="text-[10px] opacity-75 font-mono">({spk.interventionsCount})</span>
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* Recherche textuelle & Onglets d'affichage */}
          <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2.5">
            <div className="relative flex-1">
              <Search className="w-3.5 h-3.5 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Rechercher dans les interventions..."
                className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-8 pr-3 py-1.5 text-xs text-white placeholder-slate-500 focus:outline-none focus:border-blue-500 font-sans"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-500 hover:text-white text-xs"
                >
                  ✕
                </button>
              )}
            </div>

            <div className="flex items-center gap-1 bg-slate-950 p-1 rounded-lg border border-slate-800 shrink-0">
              <button
                onClick={() => setActiveTab('cues')}
                className={`px-3 py-1 text-xs font-semibold rounded-md transition-all ${
                  activeTab === 'cues'
                    ? 'bg-blue-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                Vue Segmentée ({filteredCues.length})
              </button>
              <button
                onClick={() => setActiveTab('raw')}
                className={`px-3 py-1 text-xs font-semibold rounded-md transition-all flex items-center gap-1.5 ${
                  activeTab === 'raw'
                    ? 'bg-blue-600 text-white shadow-sm'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                <Code className="w-3 h-3" />
                Format Brut .vtt
              </button>
            </div>
          </div>
        </div>

        {/* 4. Corps de la transcription (Timeline ou Brut) */}
        <div className="flex-1 overflow-y-auto p-4 sm:p-5 space-y-3 bg-slate-950/20">
          {activeTab === 'cues' ? (
            filteredCues.length === 0 ? (
              <div className="text-center py-12 text-slate-500 text-xs space-y-1">
                <FileText className="w-8 h-8 text-slate-600 mx-auto stroke-1" />
                <p className="font-semibold text-slate-400">Aucun segment ne correspond aux filtres actifs.</p>
                <p className="text-[11px]">Modifiez la recherche textuelle ou le filtre par intervenant.</p>
              </div>
            ) : (
              <div className="space-y-2.5">
                {filteredCues.map((cue) => {
                  const speakerInfo = content?.speakers?.find((s) => s.name === cue.speakerName);
                  return (
                    <div
                      key={cue.id}
                      className="p-3 bg-slate-950 rounded-xl border border-slate-800/80 hover:border-slate-700 transition-colors space-y-1.5"
                    >
                      <div className="flex items-center justify-between text-xs">
                        <div className="flex items-center gap-2">
                          <div
                            className={`w-6 h-6 rounded-md bg-gradient-to-br ${
                              speakerInfo?.avatarColor || 'from-blue-600 to-indigo-800'
                            } flex items-center justify-center text-white font-bold text-[10px] shadow-sm`}
                          >
                            {cue.speakerName.slice(0, 2).toUpperCase()}
                          </div>
                          <span className="font-semibold text-white">{cue.speakerName}</span>
                        </div>
                        <span className="font-mono text-[11px] text-slate-500 bg-slate-900 px-2 py-0.5 rounded border border-slate-800">
                          {cue.startTime} → {cue.endTime}
                        </span>
                      </div>
                      <p className="text-xs text-slate-200 leading-relaxed pl-8">
                        {cue.text}
                      </p>
                    </div>
                  );
                })}
              </div>
            )
          ) : (
            <div className="space-y-2">
              <div className="flex items-center justify-between text-[11px] text-slate-400 pb-1">
                <span>Flux textuel WebVTT standard renvoyé par Microsoft Graph API :</span>
                <span>{content?.rawContent ? `${content.rawContent.length} caractères` : '0 octet'}</span>
              </div>
              <pre className="p-3.5 bg-slate-950 rounded-xl border border-slate-800 font-mono text-[11px] text-slate-300 overflow-x-auto whitespace-pre-wrap leading-relaxed max-h-[50vh]">
                {content?.rawContent || 'Aucun contenu brut WebVTT disponible.'}
              </pre>
            </div>
          )}
        </div>

        {/* 5. Pied de page modal */}
        <div className="p-3 sm:p-4 border-t border-slate-800 bg-slate-950/80 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs">
          <div className="text-[11px] text-slate-400 font-mono flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-emerald-400"></span>
            <span>Attribution nominative des locuteurs validée</span>
          </div>

          <div className="flex items-center gap-2 w-full sm:w-auto">
            <button
              onClick={handleCopyRaw}
              className="px-3 py-1.5 text-xs font-semibold text-slate-300 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg flex items-center gap-1.5 transition-colors flex-1 sm:flex-initial justify-center"
            >
              {copied ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
              {copied ? 'Copié !' : 'Copier WebVTT'}
            </button>

            <button
              onClick={handleDownloadVtt}
              className="px-3 py-1.5 text-xs font-semibold text-slate-200 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-lg flex items-center gap-1.5 transition-colors flex-1 sm:flex-initial justify-center"
            >
              <Download className="w-3.5 h-3.5 text-blue-400" />
              Exporter (.vtt)
            </button>

            <button
              onClick={onClose}
              className="px-4 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-500 rounded-lg shadow-sm transition-all"
            >
              Fermer
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

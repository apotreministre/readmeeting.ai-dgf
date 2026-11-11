import React, { useState } from 'react';
import {
  FileText,
  Download,
  CheckCircle2,
  Clock,
  Building2,
  Calendar,
  Share2,
  Search,
  Eye,
  FileCheck
} from 'lucide-react';
import { OfficialReportItem } from '../types/meeting';

export const ReportsLibraryView: React.FC = () => {
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [selectedReport, setSelectedReport] = useState<OfficialReportItem | null>(null);

  const officialReports: OfficialReportItem[] = [
    {
      id: 'cr-1',
      referenceNumber: 'CR-DGF-2026-084',
      meetingTitle: 'Comité de Trésorerie et Arbitrage des Titres Publics du Trésor',
      meetingDate: '04 Mars 2026',
      departmentCode: 'DPE (Dette & Emprunts)',
      chairperson: 'M. le Directeur Général des Financements',
      rapporteur: 'Mme Bamba Salimata (Chef de Service)',
      status: 'approved',
      sharePointDocxUrl: 'https://financesgouvci.sharepoint.com/sites/dgf-documents/cr-2026-084.docx',
      executiveSummary: "Arbitrage des volumes d'émission sur le marché régional UMOA-Titres pour le premier trimestre 2026. Validation du calendrier indicatif d'émissions des Bons et Obligations du Trésor et sécurisation des engagements de service de la dette.",
      actionItemsCount: 5,
    },
    {
      id: 'cr-2',
      referenceNumber: 'CR-DGF-2026-083',
      meetingTitle: 'Revue de Portefeuille des Financements Multilatéraux (Banque Mondiale - DGF)',
      meetingDate: '26 Février 2026',
      departmentCode: 'DFEE (Financements Extérieurs)',
      chairperson: 'Directeur des Financements Extérieurs',
      rapporteur: 'M. Kouassi Jean-Marc (Chargé d Études)',
      status: 'under_review',
      sharePointDocxUrl: 'https://financesgouvci.sharepoint.com/sites/dgf-documents/cr-2026-083.docx',
      executiveSummary: "Examen des taux de décaissement des projets d'infrastructures énergétiques et numériques. Levée des conditions préalables d'entrée en vigueur pour le prêt d'appui aux réformes économiques.",
      actionItemsCount: 8,
    },
    {
      id: 'cr-3',
      referenceNumber: 'CR-DGF-2026-082',
      meetingTitle: 'Comité Technique d Examen des Conventions de Partenariat Public-Privé (PPP)',
      meetingDate: '19 Février 2026',
      departmentCode: 'DFEE / Cellule Juridique',
      chairperson: 'Directeur Général Adjoint',
      rapporteur: 'M. Touré Ibrahim (Conseiller Juridique)',
      status: 'published',
      sharePointDocxUrl: 'https://financesgouvci.sharepoint.com/sites/dgf-documents/cr-2026-082.docx',
      executiveSummary: "Validation des garanties souveraines et analyse des risques budgétaires contingents liés au projet de la nouvelle zone logistique d'Abidjan. Recommandations transmises au Cabinet du Ministre.",
      actionItemsCount: 4,
    },
  ];

  const filteredReports = officialReports.filter((r) => {
    return (
      r.referenceNumber.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.meetingTitle.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.departmentCode.toLowerCase().includes(searchQuery.toLowerCase())
    );
  });

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-16">
      {/* Top Banner */}
      <div className="bg-slate-900 border border-slate-800 rounded-xl p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-base font-bold text-white flex items-center gap-2">
            <FileCheck className="w-5 h-5 text-emerald-400" />
            Bibliothèque Documentaire des Comptes Rendus DGF
          </h2>
          <p className="text-xs text-slate-400 mt-1">
            Archivage officiel et synchronisation automatique avec SharePoint Online selon le modèle institutionnel agréé.
          </p>
        </div>

        <div className="relative w-full sm:w-72">
          <Search className="w-4 h-4 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            placeholder="Rechercher par référence, titre..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-9 pr-3 py-2 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-emerald-500"
          />
        </div>
      </div>

      {/* Reports Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
        {filteredReports.map((report) => (
          <div
            key={report.id}
            className="bg-slate-900/90 border border-slate-800 hover:border-slate-700 rounded-xl p-5 space-y-4 shadow-sm flex flex-col justify-between group transition-all"
          >
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <span className="font-mono text-xs font-bold text-emerald-400 bg-emerald-950/80 px-2 py-0.5 rounded border border-emerald-800/80">
                  {report.referenceNumber}
                </span>
                <span
                  className={`text-[10px] font-semibold px-2 py-0.5 rounded ${
                    report.status === 'published'
                      ? 'bg-blue-950 text-blue-300 border border-blue-800'
                      : report.status === 'approved'
                      ? 'bg-emerald-950 text-emerald-300 border border-emerald-800'
                      : 'bg-amber-950 text-amber-300 border border-amber-800'
                  }`}
                >
                  {report.status === 'published'
                    ? 'Publié au Cabinet'
                    : report.status === 'approved'
                    ? 'Validé DGF'
                    : 'En Relecture'}
                </span>
              </div>

              <h3 className="text-sm font-semibold text-white group-hover:text-emerald-300 transition-colors line-clamp-2">
                {report.meetingTitle}
              </h3>

              <p className="text-xs text-slate-400 line-clamp-3 leading-relaxed">
                {report.executiveSummary}
              </p>

              <div className="border-t border-slate-800/80 pt-3 space-y-1 text-xs text-slate-400">
                <div className="flex justify-between">
                  <span className="text-slate-500">Direction :</span>
                  <span className="text-slate-300 font-medium">{report.departmentCode}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Date de séance :</span>
                  <span className="font-mono text-slate-300">{report.meetingDate}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Actions arrêtées :</span>
                  <span className="font-mono text-emerald-400 font-bold">{report.actionItemsCount} tâches</span>
                </div>
              </div>
            </div>

            <div className="pt-2 flex items-center justify-between border-t border-slate-800">
              <button
                onClick={() => setSelectedReport(report)}
                className="text-xs font-medium text-emerald-400 hover:text-emerald-300 flex items-center gap-1.5 transition-colors"
              >
                <Eye className="w-3.5 h-3.5" />
                Aperçu du format officiel
              </button>

              <button
                onClick={() => alert(`Téléchargement du document Word DGF (${report.referenceNumber}.docx) conforme au modèle officiel.`)}
                className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded transition-colors"
                title="Télécharger le modèle Word officiel"
              >
                <Download className="w-4 h-4" />
              </button>
            </div>
          </div>
        ))}
      </div>

      {/* Official Word Document Preview Modal */}
      {selectedReport && (
        <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-2xl w-full p-6 space-y-5 shadow-2xl my-8">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <span className="text-xs font-bold text-emerald-400 font-mono">
                Aperçu Conforme Modèle Word DGF • {selectedReport.referenceNumber}
              </span>
              <button
                onClick={() => setSelectedReport(null)}
                className="text-slate-400 hover:text-white p-1 text-sm"
              >
                ✕
              </button>
            </div>

            {/* Official Document Sheet Preview */}
            <div className="bg-white text-slate-900 p-8 rounded-xl shadow-lg space-y-6 font-serif">
              {/* Official Header */}
              <div className="flex items-start justify-between border-b-2 border-slate-900 pb-4 text-center font-sans text-xs">
                <div className="text-left space-y-0.5">
                  <div className="font-bold text-slate-900 uppercase">Ministère des Finances et du Budget</div>
                  <div className="text-slate-700">Direction Générale des Financements</div>
                  <div className="font-semibold text-emerald-800">{selectedReport.departmentCode}</div>
                </div>
                <div className="text-right space-y-0.5">
                  <div className="font-bold uppercase tracking-wider">République de Côte d'Ivoire</div>
                  <div className="italic text-[11px] text-slate-600">Union - Discipline - Travail</div>
                  <div className="font-mono text-[11px] text-slate-700 font-semibold">{selectedReport.referenceNumber}</div>
                </div>
              </div>

              {/* Title */}
              <div className="text-center space-y-1 font-sans">
                <h3 className="text-base font-extrabold uppercase tracking-tight text-slate-900">
                  Compte Rendu Officiel de Séance
                </h3>
                <p className="text-xs text-slate-600 font-semibold">{selectedReport.meetingTitle}</p>
                <p className="text-[11px] text-slate-500 font-mono">Séance tenue le {selectedReport.meetingDate} à Abidjan</p>
              </div>

              {/* Executive Summary */}
              <div className="space-y-2 text-xs font-sans leading-relaxed text-slate-800">
                <h4 className="font-bold uppercase text-[11px] text-emerald-900 tracking-wider border-b border-slate-200 pb-1">
                  I. Synthèse Exécutive des Délibérations
                </h4>
                <p>{selectedReport.executiveSummary}</p>
              </div>

              {/* Action Plan */}
              <div className="space-y-2 text-xs font-sans">
                <h4 className="font-bold uppercase text-[11px] text-emerald-900 tracking-wider border-b border-slate-200 pb-1">
                  II. Matrice d'Actions & Décisions Actées ({selectedReport.actionItemsCount} points)
                </h4>
                <div className="bg-slate-50 border border-slate-200 rounded p-3 text-[11px] space-y-1.5 font-mono text-slate-700">
                  <div className="flex justify-between font-bold text-slate-900 border-b border-slate-200 pb-1">
                    <span>Action / Décision</span>
                    <span>Responsable</span>
                    <span>Échéance</span>
                  </div>
                  <div className="flex justify-between py-0.5">
                    <span>Finalisation des termes de référence financiers</span>
                    <span className="text-slate-900 font-semibold">{selectedReport.rapporteur}</span>
                    <span>15 Mars 2026</span>
                  </div>
                  <div className="flex justify-between py-0.5">
                    <span>Transmission au Cabinet du Ministre</span>
                    <span className="text-slate-900 font-semibold">{selectedReport.chairperson}</span>
                    <span>20 Mars 2026</span>
                  </div>
                </div>
              </div>

              {/* Signatures */}
              <div className="grid grid-cols-2 gap-8 pt-4 font-sans text-xs border-t border-slate-200">
                <div className="text-left space-y-8">
                  <span className="font-bold text-slate-700">Le Rapporteur de Séance :</span>
                  <div className="font-semibold text-slate-900">{selectedReport.rapporteur}</div>
                </div>
                <div className="text-right space-y-8">
                  <span className="font-bold text-slate-700">Le Président de Séance :</span>
                  <div className="font-semibold text-slate-900">{selectedReport.chairperson}</div>
                </div>
              </div>
            </div>

            <div className="flex items-center justify-between pt-2">
              <span className="text-xs text-slate-400">
                Synchronisé sur SharePoint Online (DGF Drive)
              </span>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => alert(`Document ${selectedReport.referenceNumber}.docx téléchargé.`)}
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all shadow-md"
                >
                  <Download className="w-3.5 h-3.5" />
                  Exporter Word (.docx)
                </button>
                <button
                  onClick={() => setSelectedReport(null)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-medium transition-colors"
                >
                  Fermer
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

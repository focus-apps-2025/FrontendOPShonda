import React, { useState, useEffect } from 'react';
import type { Question, OpsTemplateConfig } from '../types/forms';
import { Save, Download, Eye } from 'lucide-react';
import { generateOPSExcel } from '../utils/opsExcelGenerator';

interface ProcessStepData {
  sn: string;
  itemImportance: string;
  stepWhat: string;
  methodHow: string;
  frequencyWhen: string;
  standardCriteria: string;
  responsibility: string;
  equipmentMeasuring: string;
  possibleAbnormalities: string;
  reactionPlan: string;
  partName: string;
  partQty: string;
  ppeRequired: string;
  recordDocument: string;
  remarks: string;
}

interface OPSTemplateEditorProps {
  form: Question;
  onChange: (updatedForm: Question) => void;
  onSave?: () => void;
  onClose?: () => void;
}

export const OPSTemplateEditor: React.FC<OPSTemplateEditorProps> = ({ form, onChange, onSave, onClose }) => {
  const [config, setConfig] = useState<OpsTemplateConfig>(form.opsTemplateConfig || {});
  
  // Initialize default config if not present
  useEffect(() => {
    if (!form.opsTemplateConfig) {
      setConfig({
        basicInfoLabels: {
          deptSection: 'Dept. / Section :',
          lineZone: 'Line / Zone :',
          model: 'Model :',
          processStation: 'Process / Station :',
          formatNo: 'Format No. : ',
          controlNo: 'Control No. : '
        },
        rejectionHandling: 'REJECTION HANDLING :-\n\nClearly Identify Rejected / NG parts. Keep them properly with proper identification at defined Location.',
        measuringInstruments: [
          'Always use Calibrated Measuring Instruments / Gauges (Ensure Calibration status before using the same).',
          'Ensure Zero setting before use.',
          'Do Not Use Unidentified Measuring Tool / Gauges.',
          'In case of any abnormality, inform Line leader and Quality Engineer to take action for suspected NG material range.'
        ],
        processInstructions: [
          '2. Do Not Use Fallen Electrical/Functional Parts.',
          '3. Ensure Model / Variant Change.',
          '4. Report in case of part / hardware fallen inside vehicle.',
          '5. TQ Wrench Arrow Mark should be in correct direction.',
          '6. Put Fallen Hardware in Red Bin for Zone In-Charge judgement.',
          '7. Take approval from SH / HOD before changing process sequence.',
          '8. Zone In-Charge is overall responsible to ensure work is as per OPS.',
          '9. Contaminant parts should be covered properly.'
        ],
        tableHeaders: {
          sn: 'SN',
          itemImportance: 'Item Importance',
          stepWhat: 'Shtep \n(What / Activity)',
          methodHow: 'Method \n(How)',
          frequencyWhen: 'Frequency / When',
          standardCriteria: 'Standard \n(Spec. / Judgment Criteria)',
          responsibility: 'Responsibility',
          equipmentMeasuring: 'Equipment /\nMeasuring Eq.',
          possibleAbnormalities: 'Possible \nAbnormalities',
          reactionPlan: 'Reaction \nPlan',
          partName: 'Part Name \n& QTY',
          ppeRequired: 'PPEs\nrequired',
          recordDocument: 'Record /\nDocument',
          remarks: 'Remarks'
        },
        troubleTasks: [
          { sno: 1, trouble: 'Equipment Trouble / Machine Break Down', task: 'Stop The Line\nInform the Zone Leader\nWrite on card if mentioned in OPS' },
          { sno: 2, trouble: 'A Trouble You Are Responsible For', task: '' },
          { sno: 3, trouble: 'Empty Marshal Carrier ', task: '' },
          { sno: 4, trouble: 'Stock Out / Material Shortage ', task: '' },
          { sno: 5, trouble: 'A Trouble From Different Section', task: '' }
        ],
        abnormalityHandlingRoute: 'Abnormality handling route : \nIn case of any abnormality inform the Zone In-Charge\nFlow of Communication :-\nOperator ► Team Member ► Section Mgr ► As required',
        abnormalityDetailsLabel: 'Past Problem Details',
        associateSignArea: {
          title1: 'Associate Name \n& Emp. Code',
          title2: 'Sign & Date'
        }
      });
    }
  }, []);

  const handleChange = (path: string[], value: any) => {
    setConfig(prev => {
      const newConfig = { ...prev } as any;
      let current = newConfig;
      for (let i = 0; i < path.length - 1; i++) {
        if (!current[path[i]]) current[path[i]] = {};
        current = current[path[i]];
      }
      current[path[path.length - 1]] = value;
      return newConfig;
    });
  };

  const handleApply = () => {
    const updatedForm = { ...form, opsTemplateConfig: config };
    onChange(updatedForm);
    if (onSave) onSave();
  };

  const handleExport = () => {
    // Collect the dynamic process steps from the form's followUpQuestions to generate the Excel.
    // This replicates what happens in OPS Excel generation.
    // For this example, we'll construct dummy data for the preview.
    
    // In actual implementation, we would extract steps from form.sections/form.followUpQuestions
    const dummySteps: ProcessStepData[] = []; 
    // Just a placeholder to trigger download for testing
    const excelBuffer = generateOPSExcel(
      form.title || 'OPS Form',
      { deptSection: 'Dept', lineZone: 'Line', model: 'Model', processStation: 'Station' },
      dummySteps,
      config.abnormalityHandlingRoute || '',
      config.abnormalityDetailsLabel || '',
      {},
      config
    );
    
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `OPS_Template_${form.title}.xlsx`;
    link.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="fixed inset-0 bg-gray-100 z-50 flex flex-col overflow-hidden">
      {/* Header toolbar */}
      <div className="bg-white border-b border-gray-200 px-6 py-4 flex justify-between items-center shadow-sm">
        <h2 className="text-xl font-semibold text-gray-800">OPS Template Editor</h2>
        <div className="flex gap-4">
          <button onClick={onClose} className="px-4 py-2 text-gray-600 hover:text-gray-900 font-medium">Cancel</button>
          <button onClick={handleExport} className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 font-medium">
            <Download size={18} /> Export Excel Preview
          </button>
          <button onClick={handleApply} className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-medium">
            <Save size={18} /> Save Template
          </button>
        </div>
      </div>

      {/* Editor Content */}
      <div className="flex-1 overflow-auto p-8 max-w-6xl mx-auto w-full">
        <div className="bg-white shadow-xl rounded-lg border border-gray-300 p-8">
          <div className="mb-8 border-b pb-4">
            <h3 className="text-lg font-bold mb-4 text-blue-800">Top Header Information</h3>
            <div className="grid grid-cols-2 gap-6">
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Dept / Section Label</label>
                <input className="w-full border p-2 rounded" value={config.basicInfoLabels?.deptSection || ''} onChange={(e) => handleChange(['basicInfoLabels', 'deptSection'], e.target.value)} />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Line / Zone Label</label>
                <input className="w-full border p-2 rounded" value={config.basicInfoLabels?.lineZone || ''} onChange={(e) => handleChange(['basicInfoLabels', 'lineZone'], e.target.value)} />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Model Label</label>
                <input className="w-full border p-2 rounded" value={config.basicInfoLabels?.model || ''} onChange={(e) => handleChange(['basicInfoLabels', 'model'], e.target.value)} />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Process / Station Label</label>
                <input className="w-full border p-2 rounded" value={config.basicInfoLabels?.processStation || ''} onChange={(e) => handleChange(['basicInfoLabels', 'processStation'], e.target.value)} />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Format No Label</label>
                <input className="w-full border p-2 rounded" value={config.basicInfoLabels?.formatNo || ''} onChange={(e) => handleChange(['basicInfoLabels', 'formatNo'], e.target.value)} />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Control No Label</label>
                <input className="w-full border p-2 rounded" value={config.basicInfoLabels?.controlNo || ''} onChange={(e) => handleChange(['basicInfoLabels', 'controlNo'], e.target.value)} />
              </div>
            </div>
          </div>

          <div className="mb-8 border-b pb-4">
            <h3 className="text-lg font-bold mb-4 text-blue-800">Guidelines & Instructions</h3>
            
            <div className="mb-4">
              <label className="block text-xs font-semibold text-gray-500 mb-1">Rejection Handling</label>
              <textarea className="w-full border p-2 rounded h-24 text-sm" value={config.rejectionHandling || ''} onChange={(e) => handleChange(['rejectionHandling'], e.target.value)} />
            </div>

            <div className="mb-4">
              <label className="block text-xs font-semibold text-gray-500 mb-1">Measuring Instruments (One per line)</label>
              <textarea className="w-full border p-2 rounded h-32 text-sm" value={config.measuringInstruments?.join('\n') || ''} onChange={(e) => handleChange(['measuringInstruments'], e.target.value.split('\n'))} />
            </div>

            <div className="mb-4">
              <label className="block text-xs font-semibold text-gray-500 mb-1">Process Instructions (One per line)</label>
              <textarea className="w-full border p-2 rounded h-32 text-sm" value={config.processInstructions?.join('\n') || ''} onChange={(e) => handleChange(['processInstructions'], e.target.value.split('\n'))} />
            </div>
          </div>

          <div className="mb-8 border-b pb-4">
            <h3 className="text-lg font-bold mb-4 text-blue-800">Process Steps Table Headers</h3>
            <div className="grid grid-cols-3 gap-4 bg-gray-50 p-4 rounded border">
              {Object.entries(config.tableHeaders || {}).map(([key, value]) => (
                <div key={key}>
                  <label className="block text-xs font-semibold text-gray-500 mb-1 capitalize">{key.replace(/([A-Z])/g, ' $1')}</label>
                  <textarea className="w-full border p-2 rounded text-sm" rows={2} value={value as string} onChange={(e) => handleChange(['tableHeaders', key], e.target.value)} />
                </div>
              ))}
            </div>
          </div>

          <div className="mb-8">
            <h3 className="text-lg font-bold mb-4 text-blue-800">Footer Information</h3>
            <div className="grid grid-cols-2 gap-6">
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Abnormality Handling Route</label>
                <textarea className="w-full border p-2 rounded h-24 text-sm" value={config.abnormalityHandlingRoute || ''} onChange={(e) => handleChange(['abnormalityHandlingRoute'], e.target.value)} />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Abnormality Details Label</label>
                <input className="w-full border p-2 rounded" value={config.abnormalityDetailsLabel || ''} onChange={(e) => handleChange(['abnormalityDetailsLabel'], e.target.value)} />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Associate Sign Area (Title 1)</label>
                <textarea className="w-full border p-2 rounded h-16 text-sm" value={config.associateSignArea?.title1 || ''} onChange={(e) => handleChange(['associateSignArea', 'title1'], e.target.value)} />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 mb-1">Associate Sign Area (Title 2)</label>
                <textarea className="w-full border p-2 rounded h-16 text-sm" value={config.associateSignArea?.title2 || ''} onChange={(e) => handleChange(['associateSignArea', 'title2'], e.target.value)} />
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>
  );
};

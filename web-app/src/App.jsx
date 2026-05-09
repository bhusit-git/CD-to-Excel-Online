import { useState, useRef } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  Upload, 
  FileSpreadsheet, 
  AlertCircle, 
  CheckCircle2, 
  Folder,
  Calendar,
  Loader2,
  Download
} from 'lucide-react';
import { generateBillingWorkbook } from './lib/billingLogic';
import { parseDbf } from './lib/dbfParser';
import './App.css';

// UI Helper Components
const Card = ({ children, className = '' }) => (
  <div className={`bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden ${className}`}>
    {children}
  </div>
);

const FolderUploader = ({ label, description, onFilesSelected, selectedFiles, icon: Icon = Folder }) => {
  const fileInputRef = useRef(null);

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      onFilesSelected(Array.from(e.target.files));
    }
  };

  const dbfCount = selectedFiles.filter(f => f.name.toUpperCase().endsWith('.DBF')).length;

  return (
    <Card className="p-6 transition-all hover:border-blue-300">
      <div className="flex items-start justify-between mb-4">
        <div className="flex items-center gap-3">
          <div className="p-3 bg-blue-50 text-blue-600 rounded-lg">
            <Icon size={24} />
          </div>
          <div>
            <h3 className="font-semibold text-slate-800">{label}</h3>
            <p className="text-sm text-slate-500">{description}</p>
          </div>
        </div>
      </div>

      <div 
        onClick={() => fileInputRef.current?.click()}
        className={`mt-4 border-2 border-dashed rounded-lg p-6 flex flex-col items-center justify-center cursor-pointer transition-colors
          ${selectedFiles.length > 0 
            ? 'border-green-300 bg-green-50' 
            : 'border-slate-300 hover:border-blue-400 hover:bg-slate-50'}`}
      >
        <input
          type="file"
          webkitdirectory="true"
          directory="true"
          multiple
          className="hidden"
          ref={fileInputRef}
          onChange={handleFileChange}
        />
        
        {selectedFiles.length > 0 ? (
          <>
            <CheckCircle2 className="text-green-500 mb-2" size={32} />
            <p className="font-medium text-green-700">เลือกโฟลเดอร์แล้ว</p>
            <p className="text-xs text-green-600 mt-1">
              พบไฟล์ DBF จำนวน {dbfCount} ไฟล์ จากทั้งหมด {selectedFiles.length} ไฟล์
            </p>
          </>
        ) : (
          <>
            <Upload className="text-slate-400 mb-2" size={32} />
            <p className="font-medium text-slate-600">คลิกเพื่อเลือกโฟลเดอร์</p>
            <p className="text-xs text-slate-400 mt-1">เลือกโฟลเดอร์ที่มีไฟล์ .DBF</p>
          </>
        )}
      </div>
    </Card>
  );
};

function App() {
  const [bigSourceFiles, setBigSourceFiles] = useState([]);
  const [smallSourceFiles, setSmallSourceFiles] = useState([]);
  
  const today = new Date();
  const currentMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
  
  const [billingMonth, setBillingMonth] = useState(currentMonth);
  const [billingDate, setBillingDate] = useState(today.toISOString().split('T')[0]);
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(false);

  const handleGenerate = async () => {
    setIsProcessing(true);
    setError(null);
    setSuccess(false);
    
    try {
      // Small delay to allow UI to update to loading state
      await new Promise(r => setTimeout(r, 100));

      const loadDbfFile = async (files, fileName) => {
        const file = files.find(f => f.name.toUpperCase() === fileName.toUpperCase());
        if (!file) return [];
        const buffer = await file.arrayBuffer();
        return parseDbf(buffer);
      };

      const bigSourceData = {
        mcust: await loadDbfFile(bigSourceFiles, 'MCUST.DBF'),
        atrans: await loadDbfFile(bigSourceFiles, 'ATRANS.DBF'),
      };

      const smallSourceData = {
        mcust: await loadDbfFile(smallSourceFiles, 'MCUST.DBF'),
        abillno: await loadDbfFile(smallSourceFiles, 'ABILLNO.DBF'),
      };

      const workbookBuffer = await generateBillingWorkbook(
        bigSourceData,
        smallSourceData,
        billingMonth,
        billingDate
      );

      // Create download link
      const blob = new Blob([workbookBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Lawson_Billing_${billingMonth.replace('-', '')}.xlsx`;
      document.body.appendChild(a);
      a.click();
      
      // Cleanup
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      setSuccess(true);
    } catch (err) {
      console.error(err);
      setError(err.message || 'เกิดข้อผิดพลาดในการสร้างไฟล์ Excel');
    } finally {
      setIsProcessing(false);
    }
  };

  const isReady = bigSourceFiles.length > 0 && smallSourceFiles.length > 0 && billingMonth && billingDate;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans pb-12">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-blue-600 p-2 rounded-lg text-white">
              <FileSpreadsheet size={20} />
            </div>
            <h1 className="font-bold text-xl text-slate-900 tracking-tight">ระบบสร้างใบวางบิลลอว์สัน</h1>
          </div>
          <div className="text-sm font-medium text-slate-500 bg-slate-100 px-3 py-1.5 rounded-full">
            ประมวลผลบนเบราว์เซอร์
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 mt-8">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="mb-8"
        >
          <h2 className="text-3xl font-bold text-slate-900 mb-2">สร้างเอกสารวางบิล</h2>
          <p className="text-slate-600">อัปโหลดโฟลเดอร์ไฟล์ DBF เพื่อสร้างไฟล์รายงาน Excel สำหรับใบวางบิลลอว์สันอย่างปลอดภัย</p>
        </motion.div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          {/* Main Content Area */}
          <div className="md:col-span-2 space-y-6">
            
            <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.1 }}>
              <FolderUploader 
                label="โฟลเดอร์ต้นฉบับใหญ่ (Big Source)" 
                description="เลือกโฟลเดอร์ SI-xxxx (มีไฟล์ ATRANS.DBF และ MCUST.DBF)"
                onFilesSelected={setBigSourceFiles}
                selectedFiles={bigSourceFiles}
              />
            </motion.div>

            <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.2 }}>
              <FolderUploader 
                label="โฟลเดอร์ต้นฉบับเล็ก (Small Source)" 
                description="เลือกโฟลเดอร์ LAWSON (มีไฟล์ ABILLNO.DBF และ MCUST.DBF)"
                onFilesSelected={setSmallSourceFiles}
                selectedFiles={smallSourceFiles}
              />
            </motion.div>

          </div>

          {/* Sidebar / Configuration Area */}
          <div className="space-y-6">
            
            <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.3 }}>
              <Card className="p-6">
                <h3 className="font-semibold text-slate-800 mb-4 flex items-center gap-2">
                  <Calendar size={18} className="text-blue-500"/> ตั้งค่าการวางบิล
                </h3>
                
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">เดือนที่วางบิล</label>
                    <input 
                      type="month" 
                      value={billingMonth}
                      onChange={(e) => setBillingMonth(e.target.value)}
                      className="w-full px-3 py-2 border border-slate-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">วันที่วางบิล (หัวบิล)</label>
                    <input 
                      type="date" 
                      value={billingDate}
                      onChange={(e) => setBillingDate(e.target.value)}
                      className="w-full px-3 py-2 border border-slate-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 sm:text-sm"
                    />
                  </div>
                </div>
              </Card>
            </motion.div>

            <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.4 }}>
              <Card className="p-6">
                <h3 className="font-semibold text-slate-800 mb-4">คำสั่ง</h3>
                
                <button
                  onClick={handleGenerate}
                  disabled={!isReady || isProcessing}
                  className={`w-full flex items-center justify-center gap-2 py-3 px-4 rounded-lg font-medium transition-all shadow-sm
                    ${isReady && !isProcessing
                      ? 'bg-blue-600 hover:bg-blue-700 text-white hover:shadow-md' 
                      : 'bg-slate-100 text-slate-400 cursor-not-allowed'}`}
                >
                  {isProcessing ? (
                    <>
                      <Loader2 className="animate-spin" size={20} />
                      กำลังประมวลผล...
                    </>
                  ) : (
                    <>
                      <Download size={20} />
                      สร้างไฟล์ Excel
                    </>
                  )}
                </button>

                <AnimatePresence>
                  {error && (
                    <motion.div 
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      exit={{ opacity: 0, height: 0 }}
                      className="mt-4 p-3 bg-red-50 text-red-700 text-sm rounded-md border border-red-100 flex items-start gap-2"
                    >
                      <AlertCircle className="shrink-0 mt-0.5" size={16} />
                      <p className="break-words">{error}</p>
                    </motion.div>
                  )}
                  
                  {success && (
                    <motion.div 
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      exit={{ opacity: 0, height: 0 }}
                      className="mt-4 p-3 bg-green-50 text-green-700 text-sm rounded-md border border-green-100 flex items-start gap-2"
                    >
                      <CheckCircle2 className="shrink-0 mt-0.5" size={16} />
                      <p>สร้างไฟล์ Excel สำเร็จแล้ว!</p>
                    </motion.div>
                  )}
                </AnimatePresence>

              </Card>
            </motion.div>

          </div>
        </div>
      </main>
    </div>
  );
}

export default App;

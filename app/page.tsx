'use client';
import dynamic from 'next/dynamic';
import { useState, useEffect } from 'react';
import axios from 'axios';
import { 
  AlertCircle, 
  CheckCircle2, 
  Upload, 
  Sliders, 
  FileText, 
  CheckCircle, 
  FileCode,
  BookOpen,
  GraduationCap,
  Settings,
  Presentation,
  Monitor,
  Tv,
  Layers,
  Move,
  Maximize2,
  AlignLeft,
  X,
  Type,
  Filter
} from 'lucide-react';

function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [preset, setPreset] = useState<string>('ieee');
  const [isDragging, setIsDragging] = useState<boolean>(false);
  
  // File Type Flag
  const isPresentation = file?.name.endsWith('.ppt') || file?.name.endsWith('.pptx');

  // Excluded Pages / Slides State
  const [excludedPages, setExcludedPages] = useState<string>('');

  // Document Custom Settings State (.docx, .pdf)
  const [customTitleFont, setCustomTitleFont] = useState<string>('Times New Roman');
  const [customTitleSize, setCustomTitleSize] = useState<number>(24);
  const [customSubtitleFont, setCustomSubtitleFont] = useState<string>('Times New Roman');
  const [customSubtitleSize, setCustomSubtitleSize] = useState<number>(14);
  const [customSubtitleBold, setCustomSubtitleBold] = useState<boolean>(true);
  
  // Level 3 Sub-subheading Options (Bold & Italic)
  const [customSubsubFont, setCustomSubsubFont] = useState<string>('Times New Roman');
  const [customSubsubSize, setCustomSubsubSize] = useState<number>(12);
  const [customSubsubBold, setCustomSubsubBold] = useState<boolean>(true);
  const [customSubsubItalic, setCustomSubsubItalic] = useState<boolean>(true);

  const [customFont, setCustomFont] = useState<string>('Times New Roman');
  const [customSize, setCustomSize] = useState<number>(12);
  const [customLineSpacing, setCustomLineSpacing] = useState<number>(1.15);
  const [customMarginTop, setCustomMarginTop] = useState<number>(1.0);
  const [customMarginBottom, setCustomMarginBottom] = useState<number>(1.0);
  const [customMarginLeft, setCustomMarginLeft] = useState<number>(1.0);
  const [customMarginRight, setCustomMarginRight] = useState<number>(1.0);

  // Presentation Independent Title & Body Slide Settings
  const [aspectRatio, setAspectRatio] = useState<string>('16:9');
  
  // Title Slide Specs
  const [titleSlideFont, setTitleSlideFont] = useState<string>('Calibri');
  const [titleSlideFontSize, setTitleSlideFontSize] = useState<number>(44);

  // Body Sub-Headings Specs
  const [bodyHeadingFont, setBodyHeadingFont] = useState<string>('Times New Roman');
  const [bodyHeadingFontSize, setBodyHeadingFontSize] = useState<number>(24);

  // Body Slide Specs
  const [bodySlideFont, setBodySlideFont] = useState<string>('Times New Roman');
  const [bodySlideFontSize, setBodySlideFontSize] = useState<number>(20);

  // Slide Section Geometric & Alignment Audit Flags
  const [checkReadingOrder, setCheckReadingOrder] = useState<boolean>(true);
  const [checkBoundaryOverflow, setCheckBoundaryOverflow] = useState<boolean>(true);
  const [checkTitleAlignment, setCheckTitleAlignment] = useState<boolean>(true);
  
  // Alignment / Justification Audit Controls (Works for both Docs and Decks)
  const [checkTextJustification, setCheckTextJustification] = useState<boolean>(true);
  const [targetAlignment, setTargetAlignment] = useState<string>('LEFT');

  // Request & UI State
  const [loading, setLoading] = useState<boolean>(false);
  const [results, setResults] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);

  // Presets defined for regular documents (.docx, .pdf)
  const docPresets = [
    { id: 'ieee', label: 'IEEE Standard', icon: FileCode, desc: '2-Column, Times 10pt' },
    { id: 'apa7', label: 'APA 7th Edition', icon: GraduationCap, desc: '1-Column, 1" Margins' },
    { id: 'mla9', label: 'MLA 9th Edition', icon: BookOpen, desc: '1-Column, Double Space' },
    { id: 'custom', label: 'Custom Doc Standard', icon: Settings, desc: 'Manual Specs' },
  ];

  // Presets defined specifically for slide decks (.ppt, .pptx)
  const pptPresets = [
    { id: 'deck_16_9', label: 'Widescreen 16:9', icon: Monitor, desc: '16:9 Ratio, Standard Layout' },
    { id: 'academic_slide', label: 'Academic Defense', icon: Presentation, desc: 'High Contrast, Formal Specs' },
    { id: 'deck_4_3', label: 'Standard 4:3', icon: Tv, desc: 'Legacy Slide Ratio' },
    { id: 'custom_ppt', label: 'Custom Deck Specs', icon: Settings, desc: 'Independent Title/Body Fonts' },
  ];

  const currentPresets = isPresentation ? pptPresets : docPresets;

  useEffect(() => {
    if (isPresentation) {
      setPreset('deck_16_9');
    } else {
      setPreset('ieee');
    }
  }, [isPresentation]);

  const allowedExtensions = ['.docx', '.pdf', '.ppt', '.pptx'];

  const validateAndSetFile = (selectedFile: File) => {
    const ext = selectedFile.name.substring(selectedFile.name.lastIndexOf('.')).toLowerCase();
    if (allowedExtensions.includes(ext)) {
      setFile(selectedFile);
    } else {
      alert('Please upload a valid .docx, .pdf, .ppt, or .pptx file.');
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      validateAndSetFile(e.dataTransfer.files[0]);
    }
  };

  const handleUpload = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) return alert('Please select a file first.');

    setLoading(true);
    setError(null);
    setResults(null);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('preset', preset);
    formData.append('file_type', isPresentation ? 'presentation' : 'document');

    formData.append('excluded_pages', excludedPages);
    formData.append('check_text_justification', checkTextJustification.toString());
    formData.append('target_alignment', targetAlignment);

    if (isPresentation) {
      formData.append('check_reading_order', checkReadingOrder.toString());
      formData.append('check_boundary_overflow', checkBoundaryOverflow.toString());
      formData.append('check_title_alignment', checkTitleAlignment.toString());

      if (preset === 'custom_ppt') {
        formData.append('aspect_ratio', aspectRatio);
        formData.append('title_slide_font', titleSlideFont);
        formData.append('title_slide_font_size', titleSlideFontSize.toString());
        formData.append('body_heading_font', bodyHeadingFont);
        formData.append('body_heading_font_size', bodyHeadingFontSize.toString());
        formData.append('body_slide_font', bodySlideFont);
        formData.append('body_slide_font_size', bodySlideFontSize.toString());
      }
    } else if (preset === 'custom') {
      formData.append('custom_title_font', customTitleFont);
      formData.append('custom_title_size', customTitleSize.toString());
      formData.append('custom_subtitle_font', customSubtitleFont);
      formData.append('custom_subtitle_size', customSubtitleSize.toString());
      formData.append('custom_subtitle_bold', customSubtitleBold.toString());

      // Level 3 Heading (H3) Payload
      formData.append('custom_subsub_font', customSubsubFont);
      formData.append('custom_subsub_size', customSubsubSize.toString());
      formData.append('custom_subsub_bold', customSubsubBold.toString());
      formData.append('custom_subsub_italic', customSubsubItalic.toString());

      formData.append('custom_font', customFont);
      formData.append('custom_size', customSize.toString());
      formData.append('custom_line_spacing', customLineSpacing.toString());
      formData.append('custom_margin_top', customMarginTop.toString());
      formData.append('custom_margin_bottom', customMarginBottom.toString());
      formData.append('custom_margin_left', customMarginLeft.toString());
      formData.append('custom_margin_right', customMarginRight.toString());
    }

    try {
  const response = await axios.post('/api/verify', formData);
  setResults(response.data);
} catch (err: any) {
  setError(err.response?.data?.detail || 'An error occurred during verification.');
} finally {
  setLoading(false);
}

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800">
      
      {/* Top Navbar */}
      <header className="w-full bg-slate-900 text-white border-b border-slate-800 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 md:px-8 py-3.5 flex items-center justify-between">
          <div className="flex items-center space-x-2.5">
            <div className="p-1.5 bg-blue-600 rounded-lg">
              <FileText className="w-5 h-5 text-white" />
            </div>
            <span className="font-bold text-lg tracking-tight text-white">File Doctor</span>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <div className="max-w-7xl mx-auto p-4 md:p-8 space-y-6">
        
        {/* Header Title */}
        <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm">
          <h1 className="text-2xl font-bold text-slate-900 tracking-tight">Format Compliance Checker</h1>
          <p className="text-xs text-slate-500 mt-1">Audit document title, subheadings, typography, page margins, independent slide font styles, and text justification.</p>
        </div>

        <div className="space-y-6">
          <div className="bg-white p-6 md:p-8 rounded-xl shadow-sm border border-slate-200">
            <form onSubmit={handleUpload} className="space-y-6">
              
              {/* 1. Upload Region */}
              <div>
                <label className="block text-sm font-semibold mb-2 text-slate-900">Upload File (.docx, .pdf, .ppt, .pptx)</label>
                <div
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  onClick={() => document.getElementById('file-upload-input')?.click()}
                  className={`border-2 border-dashed rounded-xl p-8 md:p-12 text-center cursor-pointer transition-all duration-200 flex flex-col items-center justify-center min-h-[200px] ${
                    isDragging
                      ? 'border-blue-500 bg-blue-50/70 scale-[0.995]'
                      : file
                      ? 'border-emerald-500 bg-emerald-50/40'
                      : 'border-slate-300 bg-slate-50/60 hover:bg-slate-100/70 hover:border-slate-400'
                  }`}
                >
                  <input 
                    id="file-upload-input"
                    type="file" 
                    accept=".docx,.pdf,.ppt,.pptx"
                    className="hidden"
                    onChange={(e) => e.target.files?.[0] && validateAndSetFile(e.target.files[0])}
                  />

                  {file ? (
                    <div className="flex items-center justify-between w-full max-w-md bg-white p-4 rounded-xl border border-emerald-200 shadow-sm">
                      <div className="flex items-center space-x-3.5 min-w-0">
                        <div className="p-3 bg-emerald-100 rounded-lg text-emerald-600 flex-shrink-0">
                          {isPresentation ? <Presentation className="w-6 h-6" /> : <FileText className="w-6 h-6" />}
                        </div>
                        <div className="text-left min-w-0">
                          <p className="text-sm font-semibold text-slate-900 truncate">{file.name}</p>
                          <p className="text-xs text-slate-500 font-medium">
                            {(file.size / 1024).toFixed(1)} KB • {isPresentation ? 'Slide Deck' : 'Document'}
                          </p>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          setFile(null);
                        }}
                        className="p-2 hover:bg-slate-100 rounded-lg text-slate-400 hover:text-slate-700 transition flex-shrink-0"
                      >
                        <X className="w-5 h-5" />
                      </button>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      <div className="p-3.5 bg-blue-50 text-blue-600 rounded-full w-fit mx-auto shadow-xs">
                        <Upload className="w-7 h-7" />
                      </div>
                      <div>
                        <p className="text-sm font-semibold text-slate-700">
                          <span className="text-blue-600 underline underline-offset-2">Click to upload</span> or drag and drop your file here
                        </p>
                        <p className="text-xs text-slate-400 mt-1">Supports Word (.docx), PDF (.pdf), and PowerPoint (.ppt, .pptx) up to 25MB</p>
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {/* 2. Format Preset Buttons */}
              <div>
                <div className="flex items-center justify-between mb-3">
                  <label className="text-sm font-semibold text-slate-900">
                    Target {isPresentation ? 'Slide Presentation Standard' : 'Document Formatting Standard'}
                  </label>
                  {isPresentation && (
                    <span className="text-xs font-semibold px-2.5 py-1 bg-amber-100 text-amber-800 rounded-md border border-amber-200 flex items-center gap-1">
                      <Layers className="w-3.5 h-3.5" /> Slide Rules Active
                    </span>
                  )}
                </div>
                
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5">
                  {currentPresets.map((item) => {
                    const Icon = item.icon;
                    const isActive = preset === item.id;
                    return (
                      <button
                        key={item.id}
                        type="button"
                        onClick={() => setPreset(item.id)}
                        className={`p-4 rounded-xl text-left transform transition-all duration-200 ease-in-out flex flex-col justify-between border cursor-pointer hover:-translate-y-1 hover:scale-[1.02] hover:shadow-xl hover:shadow-blue-600/25 ${
                          isActive
                            ? 'bg-blue-700 border-blue-400 text-white ring-2 ring-blue-400 shadow-lg -translate-y-0.5'
                            : 'bg-blue-600 border-blue-500 text-white hover:bg-blue-700'
                        }`}
                      >
                        <div className="flex items-center justify-between w-full mb-3">
                          <div className={`p-2 rounded-lg ${isActive ? 'bg-blue-800' : 'bg-blue-700/60'}`}>
                            <Icon className="w-5 h-5 text-white" />
                          </div>
                          <div 
                            className={`w-3 h-3 rounded-full transition-all ${
                              isActive ? 'bg-emerald-400 ring-4 ring-emerald-400/30' : 'bg-blue-400/40'
                            }`} 
                          />
                        </div>
                        <div>
                          <div className="text-sm font-bold leading-tight text-white">{item.label}</div>
                          <div className={`text-[11px] mt-1 ${isActive ? 'text-blue-100 font-medium' : 'text-blue-200'}`}>{item.desc}</div>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* 3. Excluded Pages / Slides Control */}
              <div className="p-5 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
                <div className="flex items-center space-x-2 text-slate-700 font-semibold text-sm">
                  <Filter className="w-4 h-4 text-blue-600" />
                  <span>Exclude Pages / Slides</span>
                </div>
                <div className="space-y-1.5">
                  <label className="block text-xs font-semibold text-slate-600">
                    Page / Slide Numbers to Exclude
                  </label>
                  <input
                    type="text"
                    value={excludedPages}
                    onChange={(e) => setExcludedPages(e.target.value)}
                    placeholder="e.g. 1, 2, 4-6"
                    className="w-full p-2.5 border rounded-lg text-xs bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none text-slate-800 placeholder:text-slate-400 font-mono"
                  />
                  <p className="text-[11px] text-slate-400">
                    Specify comma-separated page numbers or ranges to skip predefined university templates, cover pages, or certificates during evaluation.
                  </p>
                </div>
              </div>

              {/* 4. Text Justification & Alignment Control Panel */}
              <div className="p-5 bg-slate-50 border border-slate-200 rounded-xl space-y-4">
                <div className="flex items-center space-x-2 text-slate-700 font-semibold text-sm border-b pb-2">
                  <AlignLeft className="w-4 h-4 text-blue-600" />
                  <span>Text Justification & Alignment Controls</span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <label className="flex items-center space-x-2.5 text-xs font-semibold text-slate-700 cursor-pointer bg-white p-3 rounded-lg border border-slate-200 shadow-2xs hover:bg-slate-50">
                    <input 
                      type="checkbox" 
                      checked={checkTextJustification} 
                      onChange={(e) => setCheckTextJustification(e.target.checked)}
                      className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4" 
                    />
                    <div>
                      <span className="block text-slate-900 font-bold">Enforce Text Alignment</span>
                      <span className="block text-[10px] text-slate-400 font-normal">Check paragraph text alignment</span>
                    </div>
                  </label>

                  {checkTextJustification && (
                    <div className="flex items-center justify-between bg-white p-3 rounded-lg border border-slate-200 shadow-2xs">
                      <span className="text-xs font-bold text-slate-900">Target Alignment:</span>
                      <select 
                        value={targetAlignment}
                        onChange={(e) => setTargetAlignment(e.target.value)}
                        className="text-xs p-1.5 bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-700 focus:ring-2 focus:ring-blue-500 outline-none cursor-pointer"
                      >
                        <option value="LEFT">Left Aligned</option>
                        <option value="JUSTIFY">Fully Justified</option>
                        <option value="CENTER">Centered</option>
                        <option value="RIGHT">Right Aligned</option>
                      </select>
                    </div>
                  )}
                </div>

                {/* Additional Slide-Only Controls */}
                {isPresentation && (
                  <div className="pt-3 border-t border-slate-200 space-y-3">
                    <div className="flex items-center space-x-2 text-slate-700 font-semibold text-xs">
                      <Move className="w-3.5 h-3.5 text-blue-600" />
                      <span>Presentation Slide Layout Controls</span>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                      <label className="flex items-center space-x-2.5 text-xs font-semibold text-slate-700 cursor-pointer bg-white p-3 rounded-lg border border-slate-200 shadow-2xs hover:bg-slate-50">
                        <input 
                          type="checkbox" 
                          checked={checkReadingOrder} 
                          onChange={(e) => setCheckReadingOrder(e.target.checked)}
                          className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4" 
                        />
                        <div>
                          <span className="block text-slate-900 font-bold">Vertical Order</span>
                          <span className="block text-[10px] text-slate-400 font-normal">Title top-most</span>
                        </div>
                      </label>

                      <label className="flex items-center space-x-2.5 text-xs font-semibold text-slate-700 cursor-pointer bg-white p-3 rounded-lg border border-slate-200 shadow-2xs hover:bg-slate-50">
                        <input 
                          type="checkbox" 
                          checked={checkBoundaryOverflow} 
                          onChange={(e) => setCheckBoundaryOverflow(e.target.checked)}
                          className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4" 
                        />
                        <div>
                          <span className="block text-slate-900 font-bold">Boundary Check</span>
                          <span className="block text-[10px] text-slate-400 font-normal">Canvas overflow</span>
                        </div>
                      </label>

                      <label className="flex items-center space-x-2.5 text-xs font-semibold text-slate-700 cursor-pointer bg-white p-3 rounded-lg border border-slate-200 shadow-2xs hover:bg-slate-50">
                        <input 
                          type="checkbox" 
                          checked={checkTitleAlignment} 
                          onChange={(e) => setCheckTitleAlignment(e.target.checked)}
                          className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4" 
                        />
                        <div>
                          <span className="block text-slate-900 font-bold">Header Uniformity</span>
                          <span className="block text-[10px] text-slate-400 font-normal">Header specs</span>
                        </div>
                      </label>
                    </div>
                  </div>
                )}
              </div>

              {/* 5. INDEPENDENT TITLE, SUB-HEADING & BODY SLIDE CUSTOM SETTINGS */}
              {isPresentation && preset === 'custom_ppt' && (
                <div className="p-5 bg-slate-50 border border-slate-200 rounded-xl space-y-4">
                  <div className="flex items-center justify-between border-b pb-2.5">
                    <div className="flex items-center space-x-2 text-slate-700 font-semibold text-sm">
                      <Sliders className="w-4 h-4 text-blue-600" />
                      <span>Custom Slide Deck Typography (Title, Headings vs Body)</span>
                    </div>
                    <div>
                      <select 
                        value={aspectRatio}
                        onChange={(e) => setAspectRatio(e.target.value)}
                        className="text-xs p-1.5 bg-white border rounded-md border-slate-300 font-semibold text-slate-700"
                      >
                        <option value="16:9">16:9 Widescreen</option>
                        <option value="4:3">4:3 Standard</option>
                      </select>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-3">
                      <h4 className="text-xs font-bold text-blue-900 uppercase tracking-wider flex items-center gap-1.5">
                        <Presentation className="w-4 h-4 text-blue-600" /> Title Slide Typography
                      </h4>
                      <div className="grid grid-cols-2 gap-3">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Title Font Family</label>
                          <input 
                            type="text"
                            value={titleSlideFont}
                            onChange={(e) => setTitleSlideFont(e.target.value)}
                            placeholder="e.g. Calibri"
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Title Font Size (pt)</label>
                          <input 
                            type="number"
                            value={titleSlideFontSize}
                            onChange={(e) => setTitleSlideFontSize(parseInt(e.target.value) || 0)}
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                          />
                        </div>
                      </div>
                    </div>

                    <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-3">
                      <h4 className="text-xs font-bold text-purple-900 uppercase tracking-wider flex items-center gap-1.5">
                        <Type className="w-4 h-4 text-purple-600" /> Body Sub-Headings
                      </h4>
                      <div className="grid grid-cols-2 gap-3">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Heading Font</label>
                          <input 
                            type="text"
                            value={bodyHeadingFont}
                            onChange={(e) => setBodyHeadingFont(e.target.value)}
                            placeholder="e.g. Times New Roman"
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Heading Size (pt)</label>
                          <input 
                            type="number"
                            value={bodyHeadingFontSize}
                            onChange={(e) => setBodyHeadingFontSize(parseInt(e.target.value) || 0)}
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                          />
                        </div>
                      </div>
                    </div>

                    <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-3">
                      <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
                        <FileText className="w-4 h-4 text-slate-600" /> Regular Body Text
                      </h4>
                      <div className="grid grid-cols-2 gap-3">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Body Font Family</label>
                          <input 
                            type="text"
                            value={bodySlideFont}
                            onChange={(e) => setBodySlideFont(e.target.value)}
                            placeholder="e.g. Times New Roman"
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Body Font Size (pt)</label>
                          <input 
                            type="number"
                            value={bodySlideFontSize}
                            onChange={(e) => setBodySlideFontSize(parseInt(e.target.value) || 0)}
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                          />
                        </div>
                      </div>
                    </div>

                  </div>
                </div>
              )}

              {/* 6. CUSTOM SETTINGS FOR DOCUMENTS */}
              {!isPresentation && preset === 'custom' && (
                <div className="p-5 bg-slate-50 border border-slate-200 rounded-xl space-y-4">
                  <div className="flex items-center space-x-2 text-slate-700 font-semibold text-sm border-b pb-2.5">
                    <Sliders className="w-4 h-4 text-blue-600" />
                    <span>Custom Document Typography & Margins</span>
                  </div>

                  {/* Title, Subheading (H2) & Sub-subheading (H3) Specs */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    {/* Document Title Card */}
                    <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-2.5">
                      <span className="text-xs font-bold text-slate-800 uppercase tracking-wider block">Document Title</span>
                      <div className="grid grid-cols-2 gap-2.5">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Title Font</label>
                          <input 
                            type="text"
                            value={customTitleFont}
                            onChange={(e) => setCustomTitleFont(e.target.value)}
                            placeholder="e.g. Times New Roman"
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">Title Size (pt)</label>
                          <input 
                            type="number"
                            value={customTitleSize}
                            onChange={(e) => setCustomTitleSize(parseFloat(e.target.value) || 0)}
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Subheadings / Section Headers Card (H2) */}
                    <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-2.5">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-800 uppercase tracking-wider block">Level 2 Subheading</span>
                        <label className="flex items-center space-x-1.5 text-xs font-semibold text-slate-700 cursor-pointer">
                          <input 
                            type="checkbox" 
                            checked={customSubtitleBold} 
                            onChange={(e) => setCustomSubtitleBold(e.target.checked)}
                            className="rounded text-blue-600 focus:ring-blue-500 w-3.5 h-3.5" 
                          />
                          <span>Require Bold</span>
                        </label>
                      </div>

                      <div className="grid grid-cols-2 gap-2.5">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">H2 Font</label>
                          <input 
                            type="text"
                            value={customSubtitleFont}
                            onChange={(e) => setCustomSubtitleFont(e.target.value)}
                            placeholder="e.g. Times New Roman"
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">H2 Size (pt)</label>
                          <input 
                            type="number"
                            value={customSubtitleSize}
                            onChange={(e) => setCustomSubtitleSize(parseFloat(e.target.value) || 0)}
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Level 3 Sub-subheading Card (H3) */}
                    <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-2.5">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-800 uppercase tracking-wider block">Level 3 Sub-subheading</span>
                        <div className="flex items-center space-x-2">
                          <label className="flex items-center space-x-1 text-xs font-semibold text-slate-700 cursor-pointer">
                            <input 
                              type="checkbox" 
                              checked={customSubsubBold} 
                              onChange={(e) => setCustomSubsubBold(e.target.checked)}
                              className="rounded text-blue-600 focus:ring-blue-500 w-3.5 h-3.5" 
                            />
                            <span className="font-bold">Bold</span>
                          </label>
                          <label className="flex items-center space-x-1 text-xs font-semibold text-slate-700 cursor-pointer">
                            <input 
                              type="checkbox" 
                              checked={customSubsubItalic} 
                              onChange={(e) => setCustomSubsubItalic(e.target.checked)}
                              className="rounded text-blue-600 focus:ring-blue-500 w-3.5 h-3.5" 
                            />
                            <span className="italic">Italic</span>
                          </label>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-2.5">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">H3 Font</label>
                          <input 
                            type="text"
                            value={customSubsubFont}
                            onChange={(e) => setCustomSubsubFont(e.target.value)}
                            placeholder="e.g. Times New Roman"
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">H3 Size (pt)</label>
                          <input 
                            type="number"
                            value={customSubsubSize}
                            onChange={(e) => setCustomSubsubSize(parseFloat(e.target.value) || 0)}
                            className="w-full p-2 border rounded-lg text-xs bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                          />
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Body Typography Row */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-slate-600 mb-1">Body Font Family</label>
                      <input 
                        type="text"
                        value={customFont}
                        onChange={(e) => setCustomFont(e.target.value)}
                        placeholder="e.g. Times New Roman"
                        className="w-full p-2.5 border rounded-lg text-sm bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-semibold text-slate-600 mb-1">Body Font Size (pt)</label>
                      <input 
                        type="number"
                        step="0.5"
                        value={customSize}
                        onChange={(e) => setCustomSize(parseFloat(e.target.value) || 0)}
                        className="w-full p-2.5 border rounded-lg text-sm bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-semibold text-slate-600 mb-1">Line Spacing</label>
                      <select 
                        value={customLineSpacing}
                        onChange={(e) => setCustomLineSpacing(parseFloat(e.target.value))}
                        className="w-full p-2.5 border rounded-lg text-sm bg-white border-slate-300 outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        <option value={1.0}>1.0 (Single)</option>
                        <option value={1.15}>1.15 (Standard)</option>
                        <option value={1.5}>1.5 Spaced</option>
                        <option value={2.0}>2.0 (Double)</option>
                      </select>
                    </div>
                  </div>

                  {/* Margins Row */}
                  <div>
                    <label className="block text-xs font-semibold text-slate-600 mb-2">Margin Dimensions (Inches)</label>
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                      <div>
                        <span className="text-[11px] text-slate-500 block mb-1">Top</span>
                        <input 
                          type="number"
                          step="0.05"
                          value={customMarginTop}
                          onChange={(e) => setCustomMarginTop(parseFloat(e.target.value) || 0)}
                          className="w-full p-2 border rounded-lg text-sm bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                        />
                      </div>
                      <div>
                        <span className="text-[11px] text-slate-500 block mb-1">Bottom</span>
                        <input 
                          type="number"
                          step="0.05"
                          value={customMarginBottom}
                          onChange={(e) => setCustomMarginBottom(parseFloat(e.target.value) || 0)}
                          className="w-full p-2 border rounded-lg text-sm bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                        />
                      </div>
                      <div>
                        <span className="text-[11px] text-slate-500 block mb-1">Left</span>
                        <input 
                          type="number"
                          step="0.05"
                          value={customMarginLeft}
                          onChange={(e) => setCustomMarginLeft(parseFloat(e.target.value) || 0)}
                          className="w-full p-2 border rounded-lg text-sm bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                        />
                      </div>
                      <div>
                        <span className="text-[11px] text-slate-500 block mb-1">Right</span>
                        <input 
                          type="number"
                          step="0.05"
                          value={customMarginRight}
                          onChange={(e) => setCustomMarginRight(parseFloat(e.target.value) || 0)}
                          className="w-full p-2 border rounded-lg text-sm bg-white border-slate-300 focus:ring-2 focus:ring-blue-500 outline-none"
                        />
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Submit Button */}
              <button 
                type="submit" 
                disabled={loading || !file}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3.5 rounded-xl flex items-center justify-center space-x-2 disabled:opacity-50 transition shadow-sm cursor-pointer"
              >
                <Upload className="w-5 h-5" />
                <span>{loading ? 'Auditing Format & Justification...' : 'Run Compliance Check'}</span>
              </button>
            </form>
          </div>

          {/* Bottom Panel: Audit Report & Results */}
          <div className="space-y-6">
            {error && (
              <div className="bg-red-50 text-red-700 p-4 rounded-xl flex items-center space-x-3 border border-red-200">
                <AlertCircle className="w-5 h-5 flex-shrink-0" />
                <span className="text-sm font-medium">{error}</span>
              </div>
            )}

            {!results && !error && (
              <div className="bg-white p-12 rounded-xl border border-slate-200 text-center space-y-3">
                <FileText className="w-12 h-12 text-slate-300 mx-auto" />
                <h3 className="text-base font-semibold text-slate-700">No Document Uploaded</h3>
                <p className="text-xs text-slate-400 max-w-sm mx-auto">
                  Select a target standard above, drag and drop a file, and click run to view format audit details.
                </p>
              </div>
            )}

            {results && (
              <div className="bg-white p-6 md:p-8 rounded-xl shadow-sm border border-slate-200 space-y-6">
                <div className="flex items-center justify-between border-b pb-4">
                  <div>
                    <h2 className="text-xl font-bold text-slate-900">{results.preset_name} Results</h2>
                    <p className="text-xs text-slate-500">File verified successfully</p>
                  </div>
                  <div className="text-right">
                    <span className="text-3xl font-black text-blue-600">{results.compliance_score}%</span>
                    <p className="text-xs text-slate-400">Compliance Score</p>
                  </div>
                </div>

                {/* Presentation Slide Structure Breakdown */}
                {results.slide_analysis && results.slide_analysis.length > 0 && (
                  <div className="space-y-3 border border-slate-200 p-4 rounded-xl bg-slate-50">
                    <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                      <Maximize2 className="w-4 h-4 text-blue-600" />
                      <span>Slide Geometry & Paragraph Justification Audit</span>
                    </h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                      {results.slide_analysis.map((slide: any) => (
                        <div key={slide.slide_number} className="bg-white p-3.5 rounded-lg border border-slate-200 text-xs space-y-2">
                          <div className="flex items-center justify-between border-b pb-1.5 font-bold text-slate-800">
                            <span>Slide #{slide.slide_number} ({slide.is_title_slide ? 'Title Slide' : 'Body Slide'})</span>
                            <span className={`px-2 py-0.5 rounded text-[10px] ${slide.in_order ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'}`}>
                              {slide.in_order ? 'Order Valid' : 'Order Misaligned'}
                            </span>
                          </div>
                          <div className="space-y-1">
                            {slide.sections?.map((sec: any, sIdx: number) => (
                              <div key={sIdx} className="flex justify-between text-[11px] text-slate-600 bg-slate-50 p-1.5 rounded">
                                <span className="font-semibold truncate max-w-[130px]">{sec.name}</span>
                                <span className="text-slate-400 font-mono text-[10px]">Align: {sec.alignment || 'Default'}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {/* Errors */}
                  <div className="space-y-3">
                    <h3 className="text-sm font-bold text-slate-700 flex items-center space-x-2">
                      <AlertCircle className="w-4 h-4 text-amber-500" />
                      <span>Formatting Discrepancies ({results.errors?.length || 0})</span>
                    </h3>
                    {!results.errors || results.errors.length === 0 ? (
                      <p className="text-xs text-slate-500 italic bg-slate-50 p-4 rounded-lg border border-slate-100">
                        No formatting issues detected.
                      </p>
                    ) : (
                      <div className="max-h-[350px] overflow-y-auto space-y-2.5 pr-1">
                        {results.errors.map((err: string, idx: number) => (
                          <div key={idx} className="bg-amber-50 text-amber-900 border border-amber-200 text-xs p-3.5 rounded-lg leading-relaxed">
                            {err}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Compliant Elements */}
                  <div className="space-y-3">
                    <h3 className="text-sm font-bold text-slate-700 flex items-center space-x-2">
                      <CheckCircle2 className="w-4 h-4 text-emerald-500" />
                      <span>Compliant Elements ({results.passed?.length || 0})</span>
                    </h3>
                    <div className="max-h-[350px] overflow-y-auto space-y-2.5 pr-1">
                      {results.passed?.map((pass: string, idx: number) => (
                        <div key={idx} className="bg-emerald-50 text-emerald-900 border border-emerald-200 text-xs p-3.5 rounded-lg flex items-center space-x-2">
                          <CheckCircle className="w-3.5 h-3.5 text-emerald-600 flex-shrink-0" />
                          <span>{pass}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

export default dynamic(() => Promise.resolve(Home), { ssr: false });
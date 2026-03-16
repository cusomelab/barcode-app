import React, { useState, useRef } from 'react';
import { LabelData, FIXED_DATA, CareSymbol } from '../../types/hangul';
import { CARE_SYMBOLS } from './careSymbols';

export const HangulTab: React.FC = () => {
  const [labelData, setLabelData] = useState<LabelData>({
    productName: '',
    composition: '겉감: 면 100%',
    size: 'Free',
    selectedSymbolIds: ['wash_hand_30_neutral', 'bleach_no', 'iron_140_160', 'dry_clean_petroleum'],
    customCaution: '본 제품은 소재 특성상 마찰에 의해 보풀이 일어날 수 있으니 주의하십시오.',
  });

  const printRef = useRef<HTMLDivElement>(null);

  const handleChange = (field: keyof LabelData, value: string) => {
    setLabelData(prev => ({ ...prev, [field]: value }));
  };

  const toggleSymbol = (id: string) => {
    const isSelected = labelData.selectedSymbolIds.includes(id);
    setLabelData(prev => ({
      ...prev,
      selectedSymbolIds: isSelected
        ? prev.selectedSymbolIds.filter(sid => sid !== id)
        : [...prev.selectedSymbolIds, id],
    }));
  };

  const handleDownload = async () => {
    if (printRef.current && (window as any).html2canvas) {
      try {
        const canvas = await (window as any).html2canvas(printRef.current, {
          scale: 3,
          useCORS: true,
          backgroundColor: '#ffffff',
        });
        const image = canvas.toDataURL('image/png');
        const link = document.createElement('a');
        link.href = image;
        link.download = `한글표시사항_${labelData.productName || 'label'}.png`;
        link.click();
      } catch (e) {
        alert('이미지 생성 중 오류가 발생했습니다.');
      }
    } else {
      alert('html2canvas 라이브러리를 불러오는 중입니다. 잠시 후 다시 시도해주세요.');
    }
  };

  const selectedSymbols = CARE_SYMBOLS.filter(s => labelData.selectedSymbolIds.includes(s.id));

  return (
    <div className="p-6 max-w-6xl mx-auto">
      <div className="mb-6">
        <h2 className="text-2xl font-black text-gray-900 mb-1">🏷️ 한글표시사항 생성</h2>
        <p className="text-gray-500 text-sm">의류 케어라벨(5×7cm)을 자동 생성하고 PNG로 다운로드하세요</p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* ── 입력 폼 ── */}
        <div className="space-y-5">
          {/* 기본 정보 */}
          <div className="bg-white rounded-2xl border border-gray-100 shadow-sm p-5 space-y-4">
            <h3 className="font-black text-gray-800 flex items-center gap-2">📋 정보 입력</h3>
            <div>
              <label className="block text-xs font-bold text-gray-500 mb-1">상품명 (품명)</label>
              <input type="text"
                value={labelData.productName}
                onChange={e => handleChange('productName', e.target.value)}
                className="w-full px-3 py-2.5 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-400 outline-none"
                placeholder="예: 여성용 니트 티셔츠" />
            </div>
            <div>
              <label className="block text-xs font-bold text-gray-500 mb-1">호칭 (치수)</label>
              <input type="text"
                value={labelData.size}
                onChange={e => handleChange('size', e.target.value)}
                className="w-full px-3 py-2.5 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-400 outline-none"
                placeholder="예: Free, S/M/L" />
            </div>
            <div>
              <label className="block text-xs font-bold text-gray-500 mb-1">섬유의 조성 및 혼용률</label>
              <textarea
                value={labelData.composition}
                onChange={e => handleChange('composition', e.target.value)}
                rows={3}
                className="w-full px-3 py-2.5 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-400 outline-none resize-none"
                placeholder={'겉감: 면 100%\n안감: 폴리에스터 100%'} />
            </div>
            <div>
              <label className="block text-xs font-bold text-gray-500 mb-1">추가 주의사항</label>
              <textarea
                value={labelData.customCaution}
                onChange={e => handleChange('customCaution', e.target.value)}
                rows={3}
                className="w-full px-3 py-2.5 border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-400 outline-none resize-none"
                placeholder="줄바꿈하면 자동으로 목록으로 표시됩니다." />
            </div>
          </div>

          {/* 케어심볼 선택 */}
          <div className="bg-white rounded-2xl border border-gray-100 shadow-sm p-5">
            <div className="flex items-center justify-between mb-3">
              <h3 className="font-black text-gray-800">✨ 취급상 주의사항 아이콘</h3>
              <span className={`text-xs font-bold px-2 py-1 rounded-full ${
                labelData.selectedSymbolIds.length >= 4
                  ? 'bg-green-100 text-green-700'
                  : 'bg-red-100 text-red-600'
              }`}>
                {labelData.selectedSymbolIds.length}개 선택 (최소 4개)
              </span>
            </div>
            <div className="grid grid-cols-4 sm:grid-cols-5 gap-2">
              {CARE_SYMBOLS.map(symbol => {
                const isSelected = labelData.selectedSymbolIds.includes(symbol.id);
                return (
                  <button key={symbol.id} onClick={() => toggleSymbol(symbol.id)}
                    className={`flex flex-col items-center justify-center p-2 rounded-xl border-2 transition-all ${
                      isSelected
                        ? 'border-indigo-600 bg-indigo-50 text-indigo-700'
                        : 'border-gray-200 hover:border-indigo-300 text-gray-600'
                    }`}>
                    <div className="w-8 h-8 mb-1">
                      <svg viewBox="0 0 24 24" className="w-full h-full fill-current">
                        {symbol.path}
                      </svg>
                    </div>
                    <span className="text-[9px] text-center leading-tight break-keep">{symbol.label}</span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* 고정 정보 표시 */}
          <div className="bg-gray-50 rounded-xl p-4 border border-gray-100">
            <p className="text-xs font-bold text-gray-500 mb-2">📌 고정 정보 (자동 입력)</p>
            <div className="space-y-1 text-xs text-gray-600">
              <p><span className="font-bold">수입자:</span> {FIXED_DATA.importer}</p>
              <p><span className="font-bold">주소:</span> {FIXED_DATA.address}</p>
              <p><span className="font-bold">전화:</span> {FIXED_DATA.phone}</p>
              <p><span className="font-bold">제조국:</span> {FIXED_DATA.country} | <span className="font-bold">제조연월:</span> {FIXED_DATA.date}</p>
            </div>
          </div>
        </div>

        {/* ── 미리보기 ── */}
        <div className="flex flex-col items-center">
          <div className="flex items-center justify-between w-full mb-4">
            <h3 className="font-black text-gray-800">👁️ 미리보기</h3>
            <button onClick={handleDownload}
              className="flex items-center gap-2 px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold text-sm shadow-lg transition-all">
              ⬇️ PNG 다운로드
            </button>
          </div>

          {/* 라벨 미리보기 */}
          <div
            ref={printRef}
            className="bg-white text-black font-sans relative box-border shadow-xl"
            style={{ width: '450px', height: '630px', padding: '24px', border: '1px solid #e5e7eb' }}
          >
            <div className="w-full h-full flex flex-col border-2 border-black p-1 text-sm leading-tight tracking-tight">
              {/* 헤더 */}
              <div className="text-center border-b-2 border-black pb-1 mb-1">
                <h1 className="text-xl font-bold tracking-widest">한글표시사항</h1>
              </div>

              <div className="flex-grow flex flex-col gap-1">
                {/* 품명 */}
                <div className="flex border-b border-gray-400 pb-1">
                  <span className="w-24 font-bold flex-shrink-0">품&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;명</span>
                  <span className="flex-grow font-medium">{labelData.productName || '입력필요'}</span>
                </div>
                {/* 호칭 */}
                <div className="flex border-b border-gray-400 pb-1">
                  <span className="w-24 font-bold flex-shrink-0">호&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;칭</span>
                  <span className="flex-grow font-medium">{labelData.size || '입력필요'}</span>
                </div>
                {/* 섬유조성 */}
                <div className="flex border-b border-gray-400 pb-1 min-h-[60px]">
                  <span className="w-24 font-bold flex-shrink-0">섬유의 조성</span>
                  <div className="flex-grow whitespace-pre-wrap font-medium">{labelData.composition}</div>
                </div>
                {/* 제조국 */}
                <div className="flex border-b border-gray-400 pb-1">
                  <span className="w-24 font-bold flex-shrink-0">제조국명</span>
                  <span className="flex-grow font-medium">{FIXED_DATA.country}</span>
                </div>
                {/* 제조연월 */}
                <div className="flex border-b border-gray-400 pb-1">
                  <span className="w-24 font-bold flex-shrink-0">제조연월</span>
                  <span className="flex-grow font-medium">{FIXED_DATA.date}</span>
                </div>
                {/* 수입자 */}
                <div className="flex border-b border-gray-400 pb-1">
                  <span className="w-24 font-bold flex-shrink-0">수입자명</span>
                  <span className="flex-grow font-medium">{FIXED_DATA.importer}</span>
                </div>
                {/* 주소/전화 */}
                <div className="flex border-b border-gray-400 pb-1">
                  <span className="w-24 font-bold flex-shrink-0">주소 및<br/>전화번호</span>
                  <div className="flex-grow flex flex-col justify-center">
                    <span>{FIXED_DATA.address}</span>
                    <span>{FIXED_DATA.phone}</span>
                  </div>
                </div>
                {/* 케어심볼 */}
                <div className="flex-grow flex flex-col pt-1">
                  <div className="font-bold mb-1">취급상 주의사항</div>
                  <div className="flex gap-2 justify-start mb-2 flex-wrap">
                    {selectedSymbols.length > 0 ? selectedSymbols.map(symbol => (
                      <div key={symbol.id} className="flex flex-col items-center w-14">
                        <svg viewBox="0 0 24 24" className="w-12 h-12 text-black mb-1">
                          {symbol.path}
                        </svg>
                      </div>
                    )) : (
                      <span className="text-gray-400 text-xs italic">아이콘을 선택해주세요 (최소 4개)</span>
                    )}
                  </div>
                  <div className="text-xs text-gray-800 leading-snug space-y-0.5">
                    {selectedSymbols.map(s => <p key={`txt-${s.id}`}>- {s.description}</p>)}
                    {labelData.customCaution && labelData.customCaution.split('\n').map((line, idx) =>
                      line.trim() && <p key={`custom-${idx}`}>- {line.trim()}</p>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
          <p className="mt-2 text-xs text-gray-400">실제 다운로드 시 3배 고해상도로 저장됩니다</p>
        </div>
      </div>
    </div>
  );
};

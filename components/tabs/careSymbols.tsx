import React from 'react';
import { CareSymbol } from '../../types/hangul';

// Common Styles for internal text
const TEXT_BOLD = { fontFamily: "sans-serif", fontWeight: "bold" } as const;

export const CARE_SYMBOLS: CareSymbol[] = [
  // ==========================================
  // 1. 물세탁 (Washing)
  // ==========================================
  {
    id: 'wash_95',
    category: 'washing',
    label: '95℃ 세탁',
    description: '물 온도 95℃로 세탁할 수 있습니다. 세탁기, 손세탁 가능하며 삶을 수 있습니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M3 5 L5 20 H19 L21 5" strokeLinecap="round" strokeLinejoin="round"/>
        <path d="M5 9 Q12 12 19 9" />
        <text x="12" y="16" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "6px"}}>95℃</text>
      </g>
    )
  },
  {
    id: 'wash_60',
    category: 'washing',
    label: '60℃ 세탁',
    description: '물 온도 60℃로 세탁할 수 있습니다. 세탁기 및 손세탁 가능합니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M3 5 L5 20 H19 L21 5" strokeLinecap="round" strokeLinejoin="round"/>
        <path d="M5 9 Q12 12 19 9" />
        <text x="12" y="16" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "6px"}}>60℃</text>
      </g>
    )
  },
  {
    id: 'wash_40',
    category: 'washing',
    label: '40℃ 세탁',
    description: '물 온도 40℃로 세탁할 수 있습니다. 세탁기 및 손세탁 가능합니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M3 5 L5 20 H19 L21 5" strokeLinecap="round" strokeLinejoin="round"/>
        <path d="M5 9 Q12 12 19 9" />
        <text x="12" y="16" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "6px"}}>40℃</text>
      </g>
    )
  },
  {
    id: 'wash_weak_40',
    category: 'washing',
    label: '약 40℃',
    description: '물 온도 40℃로 약하게 세탁해 주세요. (세탁기, 손세탁 가능)',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M3 5 L5 20 H19 L21 5" strokeLinecap="round" strokeLinejoin="round"/>
        <path d="M5 9 Q12 12 19 9" />
        <text x="12" y="15" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "4.5px"}}>약 40℃</text>
        <line x1="5" y1="22" x2="19" y2="22" strokeWidth="1.5" />
      </g>
    )
  },
  {
    id: 'wash_weak_30_neutral',
    category: 'washing',
    label: '약 30℃ 중성',
    description: '물 온도 30℃로 중성세제를 사용하여 약하게 세탁해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M3 5 L5 20 H19 L21 5" strokeLinecap="round" strokeLinejoin="round"/>
        <path d="M5 9 Q12 12 19 9" />
        <text x="12" y="14" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>약 30℃</text>
        <text x="12" y="18" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>중성</text>
      </g>
    )
  },
  {
    id: 'wash_hand_30_neutral',
    category: 'washing',
    label: '손세탁 30℃ 중성',
    description: '물 온도 30℃로 중성세제를 사용하여 손세탁해 주세요. (세탁기 불가)',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M3 5 L5 20 H19 L21 5" strokeLinecap="round" strokeLinejoin="round"/>
        <path d="M5 9 Q12 12 19 9" />
        <text x="12" y="13" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>손세탁</text>
        <text x="12" y="16.5" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>약 30℃</text>
        <text x="12" y="19.5" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>중성</text>
      </g>
    )
  },
  {
    id: 'wash_no',
    category: 'washing',
    label: '물세탁 금지',
    description: '물세탁은 불가능합니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M3 5 L5 20 H19 L21 5" strokeLinecap="round" strokeLinejoin="round"/>
        <path d="M5 9 Q12 12 19 9" />
        <line x1="2" y1="2" x2="22" y2="22" strokeWidth="1.5" />
        <line x1="22" y1="2" x2="2" y2="22" strokeWidth="1.5" />
      </g>
    )
  },

  // ==========================================
  // 2. 표백 (Bleaching)
  // ==========================================
  {
    id: 'bleach_chlorine',
    category: 'bleaching',
    label: '염소표백',
    description: '염소계 표백제로 표백할 수 있습니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <polygon points="12,2 22,21 2,21" strokeLinejoin="round" />
        <text x="12" y="13" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>염소</text>
        <text x="12" y="17" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>표백</text>
      </g>
    )
  },
  {
    id: 'bleach_oxygen',
    category: 'bleaching',
    label: '산소표백',
    description: '산소계 표백제로 표백할 수 있습니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <polygon points="12,2 22,21 2,21" strokeLinejoin="round" />
        <text x="12" y="13" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>산소</text>
        <text x="12" y="17" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>표백</text>
      </g>
    )
  },
  {
    id: 'bleach_chlorine_oxygen',
    category: 'bleaching',
    label: '염소,산소 표백',
    description: '염소계, 산소계 표백제로 표백할 수 있습니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <polygon points="12,2 22,21 2,21" strokeLinejoin="round" />
        <text x="12" y="12" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>염소,</text>
        <text x="12" y="15" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>산소</text>
        <text x="12" y="18" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>표백</text>
      </g>
    )
  },
  {
    id: 'bleach_no',
    category: 'bleaching',
    label: '표백 금지',
    description: '염소계, 산소계 표백제로 표백할 수 없습니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <polygon points="12,2 22,21 2,21" strokeLinejoin="round" />
        <text x="12" y="12" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>염소,</text>
        <text x="12" y="15" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>산소</text>
        <text x="12" y="18" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>표백</text>
        <line x1="8" y1="21" x2="16" y2="5" strokeWidth="1.5" strokeLinecap="round" />
        <line x1="8" y1="5" x2="16" y2="21" strokeWidth="1.5" strokeLinecap="round" />
      </g>
    )
  },

  // ==========================================
  // 3. 다림질 (Ironing) - REDESIGNED
  // ==========================================
  {
    id: 'iron_180_210',
    category: 'ironing',
    label: '180~210℃',
    description: '180~210℃로 다림질해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        {/* Iron Body: Flat bottom, curvy top */}
        <path d="M2 15 L22 15 L21 11 Q21 5 12 5 Q3 5 2 11 Z" />
        {/* Handle */}
        <path d="M7 5 Q7 1 12 1 Q17 1 17 5" strokeLinecap="round" />
        <text x="12" y="12" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>180~210℃</text>
      </g>
    )
  },
  {
    id: 'iron_140_160',
    category: 'ironing',
    label: '140~160℃',
    description: '140~160℃로 다림질해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M2 15 L22 15 L21 11 Q21 5 12 5 Q3 5 2 11 Z" />
        <path d="M7 5 Q7 1 12 1 Q17 1 17 5" strokeLinecap="round" />
        <text x="12" y="12" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>140~160℃</text>
      </g>
    )
  },
  {
    id: 'iron_80_120',
    category: 'ironing',
    label: '80~120℃',
    description: '80~120℃로 다림질해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M2 15 L22 15 L21 11 Q21 5 12 5 Q3 5 2 11 Z" />
        <path d="M7 5 Q7 1 12 1 Q17 1 17 5" strokeLinecap="round" />
        <text x="12" y="12" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>80~120℃</text>
      </g>
    )
  },
  {
    id: 'iron_cloth_80_120',
    category: 'ironing',
    label: '천 덮고 80~120℃',
    description: '원단 위에 천을 덮고 80~120℃로 다림질해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M2 14 L22 14 L21 10 Q21 4 12 4 Q3 4 2 10 Z" />
        <path d="M7 4 Q7 1 12 1 Q17 1 17 4" strokeLinecap="round" />
        <text x="12" y="11" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>80~120℃</text>
        {/* Wavy line below */}
        <path d="M3 19 Q8 22 12 19 T21 19" strokeWidth="1.2" />
      </g>
    )
  },
  {
    id: 'iron_no',
    category: 'ironing',
    label: '다림질 금지',
    description: '다림질은 불가능합니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M2 15 L22 15 L21 11 Q21 5 12 5 Q3 5 2 11 Z" />
        <path d="M7 5 Q7 1 12 1 Q17 1 17 5" strokeLinecap="round" />
        <line x1="5" y1="19" x2="19" y2="3" strokeWidth="1.5" />
        <line x1="5" y1="3" x2="19" y2="19" strokeWidth="1.5" />
      </g>
    )
  },

  // ==========================================
  // 4. 드라이클리닝 (Dry Cleaning)
  // ==========================================
  {
    id: 'dry_clean',
    category: 'drycleaning',
    label: '드라이',
    description: '드라이클리닝 가능합니다. (용제는 클로로에틸렌 또는 석유계)',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <circle cx="12" cy="12" r="9" />
        <text x="12" y="13.5" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "5px"}}>드라이</text>
      </g>
    )
  },
  {
    id: 'dry_clean_petroleum',
    category: 'drycleaning',
    label: '석유계 드라이',
    description: '드라이클리닝 가능합니다. (용제는 석유계)',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <circle cx="12" cy="12" r="9" />
        <text x="12" y="11" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "4px"}}>드라이</text>
        <text x="12" y="16" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "4px"}}>석유계</text>
      </g>
    )
  },
  {
    id: 'dry_clean_no',
    category: 'drycleaning',
    label: '드라이 금지',
    description: '드라이클리닝 불가능합니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <circle cx="12" cy="12" r="9" />
        <text x="12" y="13.5" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "5px"}}>드라이</text>
        <line x1="5" y1="19" x2="19" y2="5" strokeWidth="1.5" />
        <line x1="5" y1="5" x2="19" y2="19" strokeWidth="1.5" />
      </g>
    )
  },

  // ==========================================
  // 5. 건조 (Drying)
  // ==========================================
  {
    id: 'dry_hanger_sun',
    category: 'drying',
    label: '옷걸이(햇빛)',
    description: '햇빛에서 옷걸이에 걸어 건조해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <rect x="2" y="2" width="20" height="20" />
        {/* Hanger */}
        <path d="M6 12 C6 12 8 8 12 8 C16 8 18 12 18 12" strokeLinecap="round" />
        <path d="M12 8 V6" />
        <text x="12" y="17" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>옷걸이</text>
      </g>
    )
  },
  {
    id: 'dry_hanger_shade',
    category: 'drying',
    label: '옷걸이(그늘)',
    description: '그늘에서 옷걸이에 걸어 건조해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <rect x="2" y="2" width="20" height="20" />
        {/* Hanger */}
        <path d="M6 12 C6 12 8 8 12 8 C16 8 18 12 18 12" strokeLinecap="round" />
        <path d="M12 8 V6" />
        {/* Shade lines */}
        <line x1="2" y1="2" x2="9" y2="9" />
        <line x1="5" y1="2" x2="12" y2="9" />
        <text x="12" y="17" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>옷걸이</text>
      </g>
    )
  },
  {
    id: 'dry_flat_sun',
    category: 'drying',
    label: '뉘어서(햇빛)',
    description: '햇빛에서 바닥에 뉘어서 건조해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <rect x="2" y="2" width="20" height="20" />
        {/* Flat */}
        <path d="M6 12 H18" strokeLinecap="round" />
        <text x="12" y="17" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>뉘어서</text>
      </g>
    )
  },
  {
    id: 'dry_flat_shade',
    category: 'drying',
    label: '뉘어서(그늘)',
    description: '그늘에서 바닥에 뉘어서 건조해 주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <rect x="2" y="2" width="20" height="20" />
        {/* Flat */}
        <path d="M6 12 H18" strokeLinecap="round" />
         {/* Shade lines */}
         <line x1="2" y1="2" x2="9" y2="9" />
         <line x1="5" y1="2" x2="12" y2="9" />
        <text x="12" y="17" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3.5px"}}>뉘어서</text>
      </g>
    )
  },
  {
    id: 'dry_machine',
    category: 'drying',
    label: '기계건조 가능',
    description: '세탁 후 기계 건조 가능합니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <rect x="2" y="2" width="20" height="20" />
        <circle cx="12" cy="12" r="7" />
      </g>
    )
  },
  {
    id: 'dry_machine_no',
    category: 'drying',
    label: '기계건조 금지',
    description: '세탁 후 기계 건조 불가능합니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <rect x="2" y="2" width="20" height="20" />
        <circle cx="12" cy="12" r="7" />
        <line x1="2" y1="22" x2="22" y2="2" strokeWidth="1.5" />
        <line x1="2" y1="2" x2="22" y2="22" strokeWidth="1.5" />
      </g>
    )
  },
  {
    id: 'wring_weak',
    category: 'drying',
    label: '약하게 짬',
    description: '손으로 약하게 짜주세요.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        {/* Candy wrapper shapeish for wringing */}
        <path d="M4 6 Q8 10 4 14 V18 Q8 14 12 14 Q16 14 20 18 V14 Q16 10 20 6 V2 Q16 6 12 6 Q8 6 4 2 Z" strokeLinejoin="round" />
        <text x="12" y="10.5" textAnchor="middle" stroke="none" fill="currentColor" style={{...TEXT_BOLD, fontSize: "3px"}}>약하게</text>
      </g>
    )
  },
  {
    id: 'wring_no',
    category: 'drying',
    label: '짜면 안됨',
    description: '짜면 안 됩니다.',
    path: (
      <g stroke="currentColor" strokeWidth="1.2" fill="none">
        <path d="M4 6 Q8 10 4 14 V18 Q8 14 12 14 Q16 14 20 18 V14 Q16 10 20 6 V2 Q16 6 12 6 Q8 6 4 2 Z" strokeLinejoin="round" />
        <line x1="4" y1="18" x2="20" y2="2" strokeWidth="1.5" />
        <line x1="4" y1="2" x2="20" y2="18" strokeWidth="1.5" />
      </g>
    )
  }
];
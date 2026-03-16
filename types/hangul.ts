import React from 'react';

export interface LabelData {
  productName: string;
  composition: string;
  size: string;
  selectedSymbolIds: string[];
  customCaution: string;
}

export interface CareSymbol {
  id: string;
  category: 'washing' | 'bleaching' | 'ironing' | 'drycleaning' | 'drying';
  label: string;
  description: string; // Detailed text for the label bottom area
  path: React.ReactNode; // SVG path content
}

// Fixed data structure for the prompt's requirements
export const FIXED_DATA = {
  importer: "주식회사 폰이지",
  address: "서울시 영등포구 영등포로109, 722호",
  phone: "010-3924-0084",
  country: "중국",
  date: "2026.03",
};
/*
 * @Author: LiZhiWei
 * @Date: 2025-12-30 11:22:21
 * @LastEditors: LiZhiWei
 * @LastEditTime: 2025-12-30 11:22:22
 * @Description: 
 */
export interface ParseOptions {
  /**
   * 是否跳过某些特定的解析步骤或自定义解析行为（根据具体实现而定）
   */
  [key: string]: any;
}

export interface Size {
  width: number;
  height: number;
}

export interface SlideElement {
  type: 'text' | 'image' | 'shape' | 'chart' | 'table' | 'video' | 'audio' | string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotate?: number;
  name?: string;
  [key: string]: any;
}

export interface Slide {
  fill?: {
    type: 'color' | 'grad' | 'pic' | 'none';
    value: string | any;
  };
  elements: SlideElement[];
  layoutElements?: SlideElement[];
  note?: string;
  [key: string]: any;
}

export interface ParseResult {
  slides: Slide[];
  themeColors: Record<string, string>;
  size: Size;
}

/**
 * 解析 .pptx 文件为 JSON 数据
 * @param file .pptx 文件的 ArrayBuffer 或 Blob
 * @param options 解析选项
 */
export function parse(file: ArrayBuffer | Blob, options?: ParseOptions): Promise<ParseResult>;

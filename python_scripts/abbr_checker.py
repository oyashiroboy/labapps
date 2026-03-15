#!/usr/bin/env python3
"""
略語チェッカー (Abbreviation Checker)
Word文書から略語とその使用箇所を抽出・検索するツール

使い方:
  python3 abbr_checker.py <mode> <input_docx> [options...]

モード:
  extract   - カッコ内テキストを抽出
  search    - 略語・シノニムの出現箇所を検索
"""

import sys
import json
import re
from docx import Document

VALID_MODES = {'extract', 'search'}
ABBR_RE = re.compile(r'^[A-Za-z][A-Za-z0-9\-.]{0,49}$')
MAX_FULL_NAME_LENGTH = 200
MAX_SYNONYM_LENGTH = 80
MAX_ITEMS = 120


def extract_text_with_positions(doc):
    """
    Word文書からテキストを抽出し、段落ごとの位置情報を保持
    Introduction を境界として section を分割
    """
    paragraphs = []
    full_text_lines = []
    intro_found = False
    intro_para_idx = None
    
    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        full_text_lines.append(text)
        
        # Introduction の検出（見出しスタイル or テキスト一致）
        is_intro_heading = False
        if para.style and para.style.name:
            style_name = para.style.name.lower()
            if 'heading' in style_name:
                if re.match(r'^(1\.?\s*)?introduction\s*$', text, re.IGNORECASE):
                    is_intro_heading = True
        
        # スタイルがなくても "Introduction" だけの行は見出し扱い
        if re.match(r'^(1\.?\s*)?introduction\s*$', text, re.IGNORECASE):
            is_intro_heading = True
        
        if is_intro_heading and not intro_found:
            intro_found = True
            intro_para_idx = idx
        
        paragraphs.append({
            'index': idx,
            'text': text,
            'section': 'body' if intro_found else 'abstract'
        })
    
    return paragraphs, intro_para_idx


def extract_parentheses_content(paragraphs, max_len=20):
    """
    カッコ内のテキストを抽出（全角・半角対応）
    """
    # 半角 () と 全角 （） 両方に対応
    pattern = r'[\(（]([^\(\)（）]{1,' + str(max_len) + r'})[\)）]'
    
    candidates = []
    seen = set()
    
    for para in paragraphs:
        text = para['text']
        for match in re.finditer(pattern, text):
            content = match.group(1).strip()
            
            if content in seen:
                continue
            seen.add(content)
            
            # 略語らしさを判定
            is_likely_abbr = bool(re.match(r'^[A-Za-z][A-Za-z0-9\-\.]{0,15}$', content))
            
            # 明らかに略語じゃないものを除外
            is_unlikely = False
            # 数式・統計記号
            if re.match(r'^[np]\s*[=<>]', content, re.IGNORECASE):
                is_unlikely = True
            # 図表参照
            if re.match(r'^(Fig|Table|Eq)\.*\s*\d', content, re.IGNORECASE):
                is_unlikely = True
            # 単なる数字
            if re.match(r'^[\d\.\,\s\%\±]+$', content):
                is_unlikely = True
            
            candidates.append({
                'content': content,
                'is_likely_abbr': is_likely_abbr and not is_unlikely,
                'first_para_idx': para['index'],
                'first_section': para['section']
            })
    
    # 略語っぽいものを先に、それ以外を後に
    candidates.sort(key=lambda x: (not x['is_likely_abbr'], x['first_para_idx']))
    
    return candidates


def search_abbreviations(paragraphs, abbreviations):
    """
    略語とシノニムの出現箇所を検索
    
    abbreviations: [
        {
            'abbr': 'DPPH',
            'full_name': '2,2-Diphenyl-1-picrylhydrazyl',
            'synonyms': ['DPPH radical', 'DPPH・']
        },
        ...
    ]
    """
    results = []
    
    for abbr_info in abbreviations:
        abbr = abbr_info['abbr']
        full_name = abbr_info.get('full_name', '')
        synonyms = abbr_info.get('synonyms', [])
        
        # 検索パターンを作成（略語 + シノニム）
        search_terms = [abbr]
        if full_name:
            search_terms.append(full_name)
        search_terms.extend(synonyms)
        
        occurrences = []
        definition_found = False
        definition_para_idx = None
        
        for para in paragraphs:
            text = para['text']
            para_idx = para['index']
            section = para['section']
            
            for term in search_terms:
                # 単語境界を考慮した検索
                # 略語は大文字小文字を区別、フルネームは区別しない
                if term == abbr:
                    pattern = re.escape(term)
                    flags = 0
                else:
                    pattern = re.escape(term)
                    flags = re.IGNORECASE
                
                for match in re.finditer(pattern, text, flags):
                    start_pos = match.start()
                    end_pos = match.end()
                    
                    # 前後の文脈を取得
                    context_start = max(0, start_pos - 30)
                    context_end = min(len(text), end_pos + 30)
                    context = text[context_start:context_end]
                    if context_start > 0:
                        context = '...' + context
                    if context_end < len(text):
                        context = context + '...'
                    
                    # 定義箇所かどうか判定（カッコ内に略語がある場合）
                    is_definition = False
                    # パターン: フルネーム (略語) または 略語 (フルネーム)
                    def_pattern1 = re.escape(full_name) + r'\s*[\(（]' + re.escape(abbr) + r'[\)）]'
                    def_pattern2 = re.escape(abbr) + r'\s*[\(（]' + re.escape(full_name) + r'[\)）]'
                    
                    if full_name and (re.search(def_pattern1, text, re.IGNORECASE) or 
                                     re.search(def_pattern2, text, re.IGNORECASE)):
                        is_definition = True
                        if not definition_found:
                            definition_found = True
                            definition_para_idx = para_idx
                    
                    occurrences.append({
                        'para_idx': para_idx,
                        'section': section,
                        'term_matched': term,
                        'context': context,
                        'is_definition': is_definition
                    })
        
        # 重複を除去（同じ段落で同じ用語）
        seen_occurrences = set()
        unique_occurrences = []
        for occ in occurrences:
            key = (occ['para_idx'], occ['term_matched'])
            if key not in seen_occurrences:
                seen_occurrences.add(key)
                unique_occurrences.append(occ)
        
        # 警告チェック
        warnings = []
        
        # Abstract で定義なしに使用されている
        abstract_uses = [o for o in unique_occurrences if o['section'] == 'abstract']
        abstract_definitions = [o for o in abstract_uses if o['is_definition']]
        if abstract_uses and not abstract_definitions:
            # Abstract で定義されていない（本文で定義されている可能性）
            body_definitions = [o for o in unique_occurrences if o['section'] == 'body' and o['is_definition']]
            if body_definitions:
                warnings.append({
                    'type': 'abstract_before_definition',
                    'message': f'Abstractで使用されていますが、定義は本文（{body_definitions[0]["para_idx"]+1}段落目）にあります'
                })
        
        # 定義より前に使用されている
        if definition_para_idx is not None:
            earlier_uses = [o for o in unique_occurrences 
                          if o['para_idx'] < definition_para_idx and not o['is_definition']]
            if earlier_uses:
                warnings.append({
                    'type': 'used_before_definition',
                    'message': f'定義（{definition_para_idx+1}段落目）より前の{earlier_uses[0]["para_idx"]+1}段落目で使用されています'
                })
        
        results.append({
            'abbr': abbr,
            'full_name': full_name,
            'synonyms': synonyms,
            'occurrences': unique_occurrences,
            'definition_found': definition_found,
            'definition_para_idx': definition_para_idx,
            'warnings': warnings,
            'total_count': len(unique_occurrences)
        })
    
    return results


def main():
    if len(sys.argv) < 3:
        print(json.dumps({'error': 'Usage: abbr_checker.py <mode> <input_docx> [options...]'}))
        sys.exit(1)
    
    mode = sys.argv[1]
    input_path = sys.argv[2]

    if mode not in VALID_MODES:
        print(json.dumps({'error': f'Unknown mode: {mode}'}))
        sys.exit(1)
    
    try:
        doc = Document(input_path)
    except Exception as e:
        print(json.dumps({'error': f'Word文書を開けません: {str(e)}'}))
        sys.exit(1)
    
    # テキスト抽出
    paragraphs, intro_idx = extract_text_with_positions(doc)
    
    if mode == 'extract':
        # カッコ内テキスト抽出モード
        try:
            max_len = int(sys.argv[3]) if len(sys.argv) > 3 else 20
        except ValueError:
            print(json.dumps({'error': 'max_len must be integer'}))
            sys.exit(1)
        max_len = max(5, min(max_len, 50))
        candidates = extract_parentheses_content(paragraphs, max_len)
        
        result = {
            'mode': 'extract',
            'total_paragraphs': len(paragraphs),
            'intro_para_idx': intro_idx,
            'candidates': candidates
        }
        print(json.dumps(result, ensure_ascii=False))
    
    elif mode == 'search':
        # 検索モード
        if len(sys.argv) < 4:
            print(json.dumps({'error': 'search mode requires abbreviations JSON'}))
            sys.exit(1)
        
        try:
            abbreviations = json.loads(sys.argv[3])
        except json.JSONDecodeError as e:
            print(json.dumps({'error': f'Invalid JSON: {str(e)}'}))
            sys.exit(1)

        if not isinstance(abbreviations, list):
            print(json.dumps({'error': 'abbreviations must be an array'}))
            sys.exit(1)
        if len(abbreviations) > MAX_ITEMS:
            print(json.dumps({'error': f'too many abbreviations (max {MAX_ITEMS})'}))
            sys.exit(1)

        normalized = []
        for item in abbreviations:
            if not isinstance(item, dict):
                continue
            abbr = str(item.get('abbr', '')).strip()
            if not ABBR_RE.match(abbr):
                continue

            full_name = str(item.get('full_name', '')).strip()[:MAX_FULL_NAME_LENGTH]
            synonyms = item.get('synonyms', [])
            if not isinstance(synonyms, list):
                synonyms = []
            synonyms = [str(v).strip()[:MAX_SYNONYM_LENGTH] for v in synonyms if str(v).strip()]

            normalized.append({
                'abbr': abbr,
                'full_name': full_name,
                'synonyms': synonyms[:20]
            })

        if not normalized:
            print(json.dumps({'error': 'no valid abbreviations supplied'}))
            sys.exit(1)
        
        search_results = search_abbreviations(paragraphs, normalized)
        
        result = {
            'mode': 'search',
            'total_paragraphs': len(paragraphs),
            'intro_para_idx': intro_idx,
            'results': search_results
        }
        print(json.dumps(result, ensure_ascii=False))
    
if __name__ == '__main__':
    main()

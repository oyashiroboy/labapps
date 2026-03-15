#!/usr/bin/env python3
"""
単位・スタイルチェッカー (Unit & Style Checker)
Word文書の数値-単位間スペース、演算子、温度表記等をチェック

使い方:
  python3 unit_checker.py <input_docx> <config_json>
"""

import sys
import json
import re
from docx import Document


# プリセット定義
PRESETS = {
    'jis': {
        'name': 'JIS（日本語論文）',
        'no_space_units': ['%', '′', '″', '℃'],  # °は角度のみで使用、°Cは別途処理
        'operator_space': 'optional',
        'temperature': 'celsius',  # ℃推奨
        'liter': 'any',  # L, ℓ どちらも可
    },
    'nature': {
        'name': 'Nature系',
        'no_space_units': ['%', '′', '″'],  # °Cは温度チェックで別途処理
        'operator_space': 'required',
        'temperature': 'degree_c',  # °C必須
        'liter': 'upper',  # L必須
    },
    'acs': {
        'name': 'ACS/MDPI',
        'no_space_units': ['%', '′', '″'],  # °Cは温度チェックで別途処理
        'operator_space': 'required',
        'temperature': 'degree_c',  # °C必須
        'liter': 'upper',  # L必須
    }
}

# 時間単位パターン（フルスペルのみ検出、SI単位の min/h/s は正しいので検出しない）
TIME_PATTERNS = [
    (r'(\d+)\s*[Dd]ays?\b', 'd'),
    (r'(\d+)\s*[Hh]ours?\b', 'h'),
    (r'(\d+)\s*[Mm]inutes?\b', 'min'),
    (r'(\d+)\s*[Ss]econds?\b', 's'),
    (r'(\d+)\s*[Ww]eeks?\b', 'week'),
    (r'(\d+)\s*hrs\b', 'h'),
    (r'(\d+)\s*secs\b', 's'),
    (r'(\d+)\s*日間?', 'd'),
    (r'(\d+)\s*時間', 'h'),
    (r'(\d+)\s*分間?', 'min'),
    (r'(\d+)\s*秒間?', 's'),
]

# 既知の単位リスト（スペースが必要なもの）
KNOWN_UNITS = {
    'mg', 'g', 'kg', 'µg', 'μg', 'ng', 'pg',
    'mL', 'ml', 'L', 'l', 'µL', 'μL', 'µl', 'μl', 'dL', 'dl',
    'mol', 'mmol', 'µmol', 'μmol', 'nmol', 'pmol',
    'M', 'mM', 'µM', 'μM', 'nM', 'pM',
    'min', 'h', 's', 'd',
    'rpm', 'ppm', 'ppb',
    'kDa', 'Da', 'bp', 'kb',
    'CFU', 'IU', 'U',
    'nm', 'µm', 'μm', 'mm', 'cm', 'm', 'km',
    'Hz', 'kHz', 'MHz', 'GHz',
    'V', 'mV', 'kV', 'A', 'mA',
    'J', 'kJ', 'MJ', 'cal', 'kcal',
    'Pa', 'kPa', 'MPa', 'hPa',
    'K', 'Å', 'W', 'kW', 'MW',
    'Gy', 'Sv', 'Bq',
    'N', 'kN',
}

ALLOWED_CHECKS = {'unit_space', 'operator_space', 'temperature', 'liter', 'time_unit', 'inequality'}
ALLOWED_OPERATOR_SPACE = {'required', 'optional'}
ALLOWED_TEMPERATURE = {'celsius', 'degree_c', 'any'}
ALLOWED_LITER = {'upper', 'any'}


def extract_text_with_positions(doc):
    """Word文書からテキストを段落ごとに抽出"""
    paragraphs = []
    for idx, para in enumerate(doc.paragraphs):
        text = para.text
        if text.strip():
            paragraphs.append({
                'index': idx,
                'text': text
            })
    return paragraphs


def check_unit_space(paragraphs, config):
    """数値-単位間スペースチェック"""
    issues = []
    no_space = set(config.get('no_space_units', ['%', '′', '″']))
    
    # 大文字1文字は除外（Figure 5A などの参照を誤検出しないため）
    single_upper_exclude = set('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
    
    for para in paragraphs:
        text = para['text']
        para_idx = para['index']
        
        # °C/°Fパターンを先に見つけて除外位置を記録（スペースあり・なし両方）
        temperature_positions = set()
        for m in re.finditer(r'\d+\s*°\s*[CF]', text):
            temperature_positions.update(range(m.start(), m.end()))
        
        # 括弧内の位置を記録（型番の可能性があるので要確認扱い）
        paren_positions = set()
        paren_depth = 0
        for i, c in enumerate(text):
            if c == '(':
                paren_depth += 1
            elif c == ')':
                paren_depth -= 1
            if paren_depth > 0:
                paren_positions.add(i)
        
        # パターン1: スペースが必要なのにない（数字+英字）
        pattern = r'(\d)([a-zA-Zµμℓ][a-zA-Z]{0,4})\b'
        for match in re.finditer(pattern, text):
            # 温度表記の一部ならスキップ
            if match.start() in temperature_positions or match.end()-1 in temperature_positions:
                continue
            
            unit = match.group(2)
            full = match.group(0)
            
            # スペース不要リストにある場合はスキップ
            if unit in no_space or any(unit.startswith(u) for u in no_space):
                continue
            
            # 大文字1文字はスキップ（Figure 5A, Table 2B など）
            if unit in single_upper_exclude:
                continue
            
            # 化学式っぽいものはスキップ（C6H12O6など）
            if re.match(r'^[A-Z][a-z]?\d', full):
                continue
            
            # °C, °F などの温度表記はスキップ（温度チェックで別途処理）
            # マッチ位置の前に°があるかチェック
            if match.start() > 0 and text[match.start()-1] == '°':
                continue
            
            # 括弧内の場合は型番の可能性があるので「要確認」扱い（警告レベル）
            if match.start() in paren_positions:
                issues.append({
                    'type': 'missing_space',
                    'rule': 'unit_space',
                    'severity': 'warning',
                    'para_idx': para_idx,
                    'original': full,
                    'suggestion': f'{match.group(1)} {match.group(2)}',
                    'context': get_context(text, match.start(), match.end()),
                    'message': '数値と単位の間にスペースが必要？（型番の可能性あり、要確認）'
                })
            else:
                issues.append({
                    'type': 'missing_space',
                    'rule': 'unit_space',
                    'severity': 'error',
                    'para_idx': para_idx,
                    'original': full,
                    'suggestion': f'{match.group(1)} {match.group(2)}',
                    'context': get_context(text, match.start(), match.end()),
                    'message': '数値と単位の間にスペースが必要です'
                })
        
        # パターン2: スペース不要単位の前に不要なスペースがある
        # ただし °C/°F の場合は除外（温度表記では数字と°の間にスペースがあってもよい）
        for unit in no_space:
            if unit == '°':
                # °は温度表記で使われるので、°C/°Fの場合はスキップ
                # 純粋な角度表記（°の後にC/Fがない）のみチェック
                pattern = rf'(\d)\s+°(?![CF])'
            elif unit in ['′', '″']:
                # 分・秒記号
                pattern = rf'(\d)\s+{re.escape(unit)}'
            else:
                pattern = rf'(\d)\s+{re.escape(unit)}(?!\w)'
            
            for match in re.finditer(pattern, text):
                issues.append({
                    'type': 'extra_space',
                    'rule': 'unit_space',
                    'severity': 'error',
                    'para_idx': para_idx,
                    'original': match.group(0),
                    'suggestion': f'{match.group(1)}{unit}',
                    'context': get_context(text, match.start(), match.end()),
                    'message': f'{unit} の前にスペースは不要です'
                })
    
    return issues


def check_operator_space(paragraphs, config):
    """演算子周りのスペースチェック"""
    if config.get('operator_space') != 'required':
        return []
    
    issues = []
    # =, <, >, +, -, ×, ÷ の前後
    operators = [r'=', r'<', r'>', r'\+', r'×', r'÷']
    
    for para in paragraphs:
        text = para['text']
        para_idx = para['index']
        
        for op in operators:
            # 前後両方にスペースがない場合
            pattern = rf'(\S){op}(\S)'
            for match in re.finditer(pattern, text):
                # 例外: <=, >=, +=, -= などの複合演算子
                full = match.group(0)
                if re.match(r'[<>=!]+', full):
                    continue
                # 例外: p<0.05 などの統計表記（警告レベル下げ）
                if re.match(r'[pnr][<>=]', full, re.IGNORECASE):
                    issues.append({
                        'type': 'missing_operator_space',
                        'rule': 'operator_space',
                        'severity': 'info',
                        'para_idx': para_idx,
                        'original': full,
                        'suggestion': f'{match.group(1)} {op.replace(chr(92), "")} {match.group(2)}',
                        'context': get_context(text, match.start(), match.end()),
                        'message': '演算子の前後にスペースを入れることを推奨します'
                    })
                    continue
                
                # 演算子の実際の文字を取得（エスケープを除去）
                op_char = op.replace('\\', '')
                issues.append({
                    'type': 'missing_operator_space',
                    'rule': 'operator_space',
                    'severity': 'warning',
                    'para_idx': para_idx,
                    'original': full,
                    'suggestion': f'{match.group(1)} {op_char} {match.group(2)}',
                    'context': get_context(text, match.start(), match.end()),
                    'message': '演算子の前後にスペースが必要です'
                })
    
    return issues


def check_temperature(paragraphs, config):
    """温度表記チェック"""
    issues = []
    temp_style = config.get('temperature', 'any')
    
    celsius_pattern = r'(\d+)\s*℃'
    degree_c_pattern = r'(\d+)\s*°\s*C'
    
    celsius_count = 0
    degree_c_count = 0
    celsius_locations = []
    degree_c_locations = []
    
    for para in paragraphs:
        text = para['text']
        para_idx = para['index']
        
        for match in re.finditer(celsius_pattern, text):
            celsius_count += 1
            celsius_locations.append({
                'para_idx': para_idx,
                'text': match.group(0),
                'context': get_context(text, match.start(), match.end())
            })
        
        for match in re.finditer(degree_c_pattern, text):
            degree_c_count += 1
            degree_c_locations.append({
                'para_idx': para_idx,
                'text': match.group(0),
                'context': get_context(text, match.start(), match.end())
            })
    
    # 揺れ検出
    if celsius_count > 0 and degree_c_count > 0:
        issues.append({
            'type': 'temperature_inconsistency',
            'rule': 'temperature',
            'severity': 'warning',
            'message': f'温度表記が混在: ℃({celsius_count}箇所) と °C({degree_c_count}箇所)',
            'details': {
                'celsius': celsius_locations[:5],
                'degree_c': degree_c_locations[:5]
            }
        })
    
    # スタイル違反
    if temp_style == 'celsius' and degree_c_count > 0:
        issues.append({
            'type': 'temperature_style',
            'rule': 'temperature',
            'severity': 'info',
            'message': f'JISでは ℃ が推奨です（°C が{degree_c_count}箇所）',
            'details': {'locations': degree_c_locations[:5]}
        })
    elif temp_style == 'degree_c' and celsius_count > 0:
        issues.append({
            'type': 'temperature_style',
            'rule': 'temperature',
            'severity': 'info',
            'message': f'このスタイルでは °C が推奨です（℃ が{celsius_count}箇所）',
            'details': {'locations': celsius_locations[:5]}
        })
    
    return issues


def check_liter(paragraphs, config):
    """リットル表記チェック"""
    issues = []
    liter_style = config.get('liter', 'any')
    
    # 小文字l, 大文字L, ℓ の検出
    patterns = {
        'lower': r'\b(m|µ|μ)?l\b',
        'upper': r'\b(m|µ|μ)?L\b', 
        'symbol': r'(m|µ|μ)?ℓ'
    }
    
    counts = {'lower': 0, 'upper': 0, 'symbol': 0}
    locations = {'lower': [], 'upper': [], 'symbol': []}
    
    for para in paragraphs:
        text = para['text']
        para_idx = para['index']
        
        for key, pattern in patterns.items():
            for match in re.finditer(pattern, text):
                counts[key] += 1
                if len(locations[key]) < 5:
                    locations[key].append({
                        'para_idx': para_idx,
                        'text': match.group(0),
                        'context': get_context(text, match.start(), match.end())
                    })
    
    # 揺れ検出
    used_styles = [k for k, v in counts.items() if v > 0]
    if len(used_styles) > 1:
        msg_parts = []
        if counts['upper'] > 0:
            msg_parts.append(f'L({counts["upper"]}箇所)')
        if counts['lower'] > 0:
            msg_parts.append(f'l({counts["lower"]}箇所)')
        if counts['symbol'] > 0:
            msg_parts.append(f'ℓ({counts["symbol"]}箇所)')
        
        issues.append({
            'type': 'liter_inconsistency',
            'rule': 'liter',
            'severity': 'warning',
            'message': f'リットル表記が混在: {" と ".join(msg_parts)}',
            'details': locations
        })
    
    # スタイル違反
    if liter_style == 'upper':
        if counts['lower'] > 0 or counts['symbol'] > 0:
            issues.append({
                'type': 'liter_style',
                'rule': 'liter',
                'severity': 'info',
                'message': 'このスタイルでは大文字 L が推奨です'
            })
    
    return issues


def check_time_units(paragraphs, config):
    """時間単位チェック（SI推奨）"""
    issues = []
    
    for para in paragraphs:
        text = para['text']
        para_idx = para['index']
        
        for pattern, si_unit in TIME_PATTERNS:
            for match in re.finditer(pattern, text):
                original = match.group(0)
                number = match.group(1)
                
                issues.append({
                    'type': 'time_unit',
                    'rule': 'time_unit',
                    'severity': 'info',
                    'para_idx': para_idx,
                    'original': original,
                    'suggestion': f'{number} {si_unit}',
                    'context': get_context(text, match.start(), match.end()),
                    'message': f'SI単位では "{number} {si_unit}" が推奨です'
                })
    
    return issues


def check_inequality(paragraphs, config):
    """不等号の揺れチェック"""
    issues = []
    
    # ≧(U+2267) vs ≥(U+2265), ≦(U+2266) vs ≤(U+2264)
    inequality_chars = {
        '≧': '≥',  # 二重線 → 単線
        '≦': '≤',
        '≥': '≧',
        '≤': '≦',
    }
    
    # 各記号の出現箇所を収集
    locations = {'geq_double': [], 'geq_single': [], 'leq_double': [], 'leq_single': []}
    
    for para in paragraphs:
        text = para['text']
        para_idx = para['index']
        
        for i, char in enumerate(text):
            if char == '≧':
                locations['geq_double'].append({
                    'para_idx': para_idx,
                    'char': char,
                    'context': get_context(text, i, i+1)
                })
            elif char == '≥':
                locations['geq_single'].append({
                    'para_idx': para_idx,
                    'char': char,
                    'context': get_context(text, i, i+1)
                })
            elif char == '≦':
                locations['leq_double'].append({
                    'para_idx': para_idx,
                    'char': char,
                    'context': get_context(text, i, i+1)
                })
            elif char == '≤':
                locations['leq_single'].append({
                    'para_idx': para_idx,
                    'char': char,
                    'context': get_context(text, i, i+1)
                })
    
    geq_double = len(locations['geq_double'])
    geq_single = len(locations['geq_single'])
    leq_double = len(locations['leq_double'])
    leq_single = len(locations['leq_single'])
    
    if (geq_double > 0 and geq_single > 0) or (leq_double > 0 and leq_single > 0):
        msg_parts = []
        if geq_double > 0:
            msg_parts.append(f'≧({geq_double})')
        if geq_single > 0:
            msg_parts.append(f'≥({geq_single})')
        if leq_double > 0:
            msg_parts.append(f'≦({leq_double})')
        if leq_single > 0:
            msg_parts.append(f'≤({leq_single})')
        
        # 各箇所を個別の issue として追加
        all_locations = (locations['geq_double'] + locations['geq_single'] + 
                        locations['leq_double'] + locations['leq_single'])
        
        for loc in all_locations:
            issues.append({
                'type': 'inequality_inconsistency',
                'rule': 'inequality',
                'severity': 'warning',
                'para_idx': loc['para_idx'],
                'original': loc['char'],
                'context': loc['context'],
                'message': f'不等号スタイルが混在: {loc["char"]} (文書全体: {" と ".join(msg_parts)})'
            })
    
    return issues


def get_context(text, start, end, context_len=30):
    """マッチ箇所の前後コンテキストを取得（該当箇所をマーク）"""
    ctx_start = max(0, start - context_len)
    ctx_end = min(len(text), end + context_len)
    
    # 該当箇所の前後を取得
    matched_text = text[start:end]
    before = text[ctx_start:start]
    after = text[end:ctx_end]
    
    # → を使って検出箇所を示す
    context = f'{before}→{matched_text}←{after}'
    if ctx_start > 0:
        context = '...' + context
    if ctx_end < len(text):
        context = context + '...'
    
    return context


def main():
    if len(sys.argv) < 3:
        print(json.dumps({'error': 'Usage: unit_checker.py <input_docx> <config_json>'}))
        sys.exit(1)
    
    input_path = sys.argv[1]
    config_json = sys.argv[2]
    
    try:
        config = json.loads(config_json)
    except json.JSONDecodeError as e:
        print(json.dumps({'error': f'Invalid config JSON: {str(e)}'}))
        sys.exit(1)

    if not isinstance(config, dict):
        print(json.dumps({'error': 'config must be an object'}))
        sys.exit(1)

    checks = config.get('checks', ['unit_space'])
    if not isinstance(checks, list):
        print(json.dumps({'error': 'checks must be an array'}))
        sys.exit(1)
    checks = [str(v) for v in checks][:10]
    invalid_checks = [c for c in checks if c not in ALLOWED_CHECKS]
    if invalid_checks:
        print(json.dumps({'error': f'invalid checks: {", ".join(invalid_checks)}'}))
        sys.exit(1)
    if not checks:
        checks = ['unit_space']

    no_space_units = config.get('no_space_units', ['%', '′', '″'])
    if not isinstance(no_space_units, list):
        print(json.dumps({'error': 'no_space_units must be an array'}))
        sys.exit(1)
    no_space_units = [str(v).strip()[:10] for v in no_space_units if str(v).strip()][:60]

    operator_space = str(config.get('operator_space', 'required'))
    if operator_space not in ALLOWED_OPERATOR_SPACE:
        print(json.dumps({'error': 'invalid operator_space'}))
        sys.exit(1)

    temperature = str(config.get('temperature', 'degree_c'))
    if temperature not in ALLOWED_TEMPERATURE:
        print(json.dumps({'error': 'invalid temperature'}))
        sys.exit(1)

    liter = str(config.get('liter', 'upper'))
    if liter not in ALLOWED_LITER:
        print(json.dumps({'error': 'invalid liter'}))
        sys.exit(1)

    config = {
        'checks': checks,
        'no_space_units': no_space_units,
        'operator_space': operator_space,
        'temperature': temperature,
        'liter': liter
    }
    
    try:
        doc = Document(input_path)
    except Exception as e:
        print(json.dumps({'error': f'Word文書を開けません: {str(e)}'}))
        sys.exit(1)
    
    paragraphs = extract_text_with_positions(doc)
    
    all_issues = []
    
    # 有効なチェックを実行
    checks = config.get('checks', ['unit_space'])
    
    if 'unit_space' in checks:
        all_issues.extend(check_unit_space(paragraphs, config))
    
    if 'operator_space' in checks:
        all_issues.extend(check_operator_space(paragraphs, config))
    
    if 'temperature' in checks:
        all_issues.extend(check_temperature(paragraphs, config))
    
    if 'liter' in checks:
        all_issues.extend(check_liter(paragraphs, config))
    
    if 'time_unit' in checks:
        all_issues.extend(check_time_units(paragraphs, config))
    
    if 'inequality' in checks:
        all_issues.extend(check_inequality(paragraphs, config))
    
    # 結果を集計
    summary = {
        'error': len([i for i in all_issues if i.get('severity') == 'error']),
        'warning': len([i for i in all_issues if i.get('severity') == 'warning']),
        'info': len([i for i in all_issues if i.get('severity') == 'info']),
    }
    
    result = {
        'total_paragraphs': len(paragraphs),
        'total_issues': len(all_issues),
        'summary': summary,
        'issues': all_issues
    }
    
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    main()

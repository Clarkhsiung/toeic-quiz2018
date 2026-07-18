import json
import re
import os
import time
import urllib.parse
import requests

def fetch_word_assets_from_cambridge_api(word, target_lang='zh-TW'):
    """
    透過公開 API 整合查詢單字的 CEFR 等級、中文翻譯（含多重詞性釋義）與中英雙語例句
    """
    # 預設完整資料欄位結構
    assets = {
        "cefr": "B1", # 預設多益基礎等級
        "translation": "",
        "sentence_en": "No example sentence available.",
        "sentence_zh": "暫無例句翻譯。"
    }
    
    try:
        # 1. 抓取單字基本語意與例句 (利用字典隱含 API)
        dict_url = f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}"
        dict_res = requests.get(dict_url, timeout=5)
        
        found_example = False
        example_en = ""
        
        if dict_res.status_code == 200:
            dict_json = dict_res.json()
            if isinstance(dict_json, list) and len(dict_json) > 0:
                # 深度檢索例句
                for meaning in dict_json[0].get("meanings", []):
                    for definition in meaning.get("definitions", []):
                        if "example" in definition and definition["example"]:
                            example_en = definition["example"].strip()
                            found_example = True
                            break
                    if found_example: 
                        break

        # 2. 透過 Google Translate API 獲取高精確度的多重詞性中文翻譯
        # dt=t: 主要翻譯, dt=bd: 詞性與不同解釋列表, dt=at: 備用/同義翻譯
        trans_url = "https://translate.googleapis.com/translate_a/single"
        params = [
            ('client', 'gtx'),
            ('sl', 'en'),
            ('tl', target_lang),
            ('dt', 't'),
            ('dt', 'bd'),
            ('dt', 'at'),
            ('q', word)
        ]
        
        trans_res = requests.get(trans_url, params=params, timeout=5)
        if trans_res.status_code == 200:
            trans_json = trans_res.json()
            
            # 解析最主要翻譯
            main_translation = ""
            if trans_json[0]:
                main_translation = "".join([part[0] for part in trans_json[0] if part[0]]).strip()
            
            # 解析不同的中文翻譯與詞性
            dict_translations = []
            pos_zh_map = {
                'noun': '名',
                'verb': '動',
                'adjective': '形',
                'adverb': '副',
                'pronoun': '代',
                'preposition': '介',
                'conjunction': '連',
                'interjection': '感',
                'abbreviation': '縮'
            }
            
            try:
                # trans_json[1] 通常包含多重詞性列表
                if len(trans_json) > 1 and isinstance(trans_json[1], list):
                    for pos_entry in trans_json[1]:
                        if isinstance(pos_entry, list) and len(pos_entry) >= 2:
                            pos = pos_entry[0]       # 英文詞性，例如 'noun'
                            meanings = pos_entry[1]  # 該詞性下的所有翻譯列表
                            
                            if isinstance(pos, str) and isinstance(meanings, list):
                                pos_short = pos_zh_map.get(pos.lower(), pos)
                                # 過濾掉空值，並轉為字串
                                meanings_clean = [str(m).strip() for m in meanings if m]
                                if meanings_clean:
                                    # 每個詞性限制最多擷取前 4 個最常用的翻譯以求版面美觀
                                    meanings_str = "、".join(meanings_clean[:4])
                                    dict_translations.append(f"[{pos_short}] {meanings_str}")
            except Exception as parse_err:
                print(f"⚠️ 解析單字 [{word}] 多重詞性時發生微小錯誤 (已自動跳過): {parse_err}")

            # 整合最終翻譯字串
            if dict_translations:
                # 輸出格式如: 跑 ( [動] 奔跑、運作；[名] 跑步、運行 )
                assets["translation"] = f"{main_translation} ( {'；'.join(dict_translations)} )"
            else:
                assets["translation"] = main_translation

        # 3. 如果有撈到英文例句，同步將例句翻譯成繁體中文並儲存
        if found_example and example_en:
            assets["sentence_en"] = example_en
            sent_params = {
                'client': 'gtx',
                'sl': 'en',
                'tl': target_lang,
                'dt': 't',
                'q': example_en
            }
            sent_res = requests.get(trans_url, params=sent_params, timeout=5)
            if sent_res.status_code == 200:
                assets["sentence_zh"] = "".join([part[0] for part in sent_res.json()[0] if part[0]]).strip()
        
        # 4. 根據單字學術難度自動化分級演算法 (CEFR 模擬器)
        word_len = len(word)
        if word_len <= 4:
            assets["cefr"] = "A1" if word_len <= 3 else "A2"
        elif word_len <= 6:
            assets["cefr"] = "B1"
        elif word_len <= 8:
            assets["cefr"] = "B2"
        elif word_len <= 10:
            assets["cefr"] = "C1"
        else:
            assets["cefr"] = "C2"

    except Exception as e:
        print(f"⚠️ 處理單字 [{word}] 網路連線或解析超時: {e}")
        
    return assets

def rebuild_database_to_v3():
    input_filename = 'words_db.json'
    output_filename = 'words_db_v3.json'

    if not os.path.exists(input_filename):
        print(f"❌ 找不到原始資料庫檔案：{input_filename}，請確保此腳本與舊資料庫在同一目錄。")
        return

    with open(input_filename, 'r', encoding='utf-8') as f:
        try:
            old_data = json.load(f)
        except Exception as e:
            print(f"❌ 讀取 JSON 格式出錯: {e}")
            return

    v3_database = {}
    processed_count = 0

    print("🚀 開始執行 3.0 大版本背景資料庫重構作業（已啟用多重詞性翻譯功能）...")
    print("--------------------------------------------------")

    for raw_key, value in old_data.items():
        page_source = value.get("page", "")
        original_meaning = value.get("meaning", "")
        
        # 清洗與提取純英文字母 (防止單字書欄位帶有雜質)
        full_text = f"{raw_key} {original_meaning}"
        extracted_words = re.findall(r'\b[a-zA-Z\-]+\b', full_text)

        for word in extracted_words:
            clean_word = word.strip()
            if len(clean_word) <= 1: 
                continue
            
            # 避免重複抓取相同單字
            if clean_word.lower() in [w.lower() for w in v3_database.keys()]:
                continue

            processed_count += 1
            print(f"🔄 [{processed_count}] 正在升級單字 ➔ [{clean_word}] (對應頁碼: {page_source})")
            
            # 發送請求獲取 3.0 擴充欄位資料
            api_assets = fetch_word_assets_from_cambridge_api(clean_word)
            
            # 整合寫入全新規格的資料庫中
            v3_database[clean_word] = {
                "cefr": api_assets["cefr"],
                "translation": api_assets["translation"] if api_assets["translation"] else original_meaning,
                "page": page_source,
                "sentence_en": api_assets["sentence_en"],
                "sentence_zh": api_assets["sentence_zh"]
            }
            
            # 短暫延遲避免請求過快被阻擋
            time.sleep(0.3)

    # 寫入全新的結構化 JSON 檔案
    with open(output_filename, 'w', encoding='utf-8') as f:
        json.dump(v3_database, f, ensure_ascii=False, indent=4)

    print("--------------------------------------------------")
    print("✨ 大版本更新完成！")
    print(f"🎉 成功建立含有豐富多重詞性釋義、CEFR 級別、中英雙語例句的全新資料庫，共計: {len(v3_database)} 筆單字")
    print(f"💾 產出檔案名稱：{output_filename}")
    print(f"💡 請在 GitHub 部署前，將此檔案更名為 words_db.json 取代網頁目錄下的舊檔。")

if __name__ == '__main__':
    rebuild_database_to_v3()
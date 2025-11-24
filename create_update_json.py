import json
import os

OUTPUT_PATH = os.path.join('005_ToolOutput', '02_ResumeUpdate', 'Data', 'resume_update.json')

data = {
  "update_payload": [
    {
      "action": "INSERT",
      "target_no": 0,
      "data": {
        "period": {
          "start": "2025/11",
          "end": "現在"
        },
        "business_content": {
          "title_col_e": [
            "製造（化粧品）既存SAP（MM/SD）改修・要件定義",
            "",
            "",
            "",
            ""
          ],
          "role_col_f": [
            "",
            "主な役割",
            "",
            "",
            ""
          ],
          "detail_col_g": [
            "",
            "",
            "既存アドオン（MM/SD）のコード解析、影響調査、見積",
            "要件定義、基本設計、技術的負債の解消（リファクタリング）",
            ""
          ]
        },
        "technology": {
          "environment_col_u": [
            "SAP ECC 6.0",
            "",
            "",
            "",
            ""
          ],
          "language_col_z": [
            "ABAP",
            "",
            "",
            "",
            ""
          ],
          "process_col_ae": [
            "アプリ開発チーム",
            "SAPコンサルタント",
            "開発リーダ",
            "",
            ""
          ]
        }
      }
    },
    {
      "action": "UPDATE",
      "target_no": 1,
      "data": {
        "period": {
          "start": "2023/4",
          "end": "2025/10"
        },
        "business_content": {
          "title_col_e": [
            "製造（医療機器）工場のSAP（FI）導入",
            "",
            "",
            "",
            ""
          ],
          "role_col_f": [
            "",
            "主な役割",
            "",
            "",
            ""
          ],
          "detail_col_g": [
            "",
            "",
            "基本設計、詳細設計、開発、開発管理、結合テスト管理",
            "システムテスト管理（IF／マスタ関連）、UAT支援、",
            "初期流動対応"
          ]
        },
        "technology": {
          "environment_col_u": [
            "S4/HANA　1709",
            "Linux",
            "",
            "",
            ""
          ],
          "language_col_z": [
            "ABAP/4",
            "",
            "",
            "",
            ""
          ],
          "process_col_ae": [
            "アプリ開発チーム",
            "開発リーダ",
            "",
            "",
            ""
          ]
        }
      }
    }
  ],
  "footer_update": {
      "update_required": True,
      "other_col_b": [
          "・FI/MM等、複数モジュールにおける要件定義から保守までの経験と、10名程度のチームリード実績があります。",
          "・開発・検証・移行の各局面において、技術的な観点からプロジェクトの円滑な進行と課題解決を支援した経験があります。"
      ]
  }
}

with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=2, ensure_ascii=False)

print(f"Generated {OUTPUT_PATH}")

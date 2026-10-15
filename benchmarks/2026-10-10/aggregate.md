| | **docx-cli 0.27.0** | **Anthropic docx skill** | **OfficeCLI 1.0.156** |
| :-- | --: | --: | --: |
| Tasks solved (of 6) | 6.0 (6.0–6.0) | 6.0 (6.0–6.0) | 4.7 (4.0–5.0) |
| Rendered correctly (of 6) | 6.0 (6.0–6.0) | 6.0 (6.0–6.0) | 6.0 (6.0–6.0) |
| Effective input tokens | 902k (899k–904k) | 1.15M (1.09M–1.22M) | 1.92M (1.78M–2.00M) |
| Output tokens | 32k (27k–35k) | 54k (48k–66k) | 63k (59k–66k) |
| Wall-clock (sum of 6 tasks) | 914 s (808 s–990 s) | 1,254 s (1,039 s–1,511 s) | 1,834 s (1,659 s–2,085 s) |
| Tool calls | 98 (98–98) | 124 (118–128) | 165 (157–174) |

| Scenario | docx-cli 0.27.0 pass@1 · pass@3 · pass^3 | Anthropic docx skill pass@1 · pass@3 · pass^3 | OfficeCLI 1.0.156 pass@1 · pass@3 · pass^3 |
| :-- | :--: | :--: | :--: |
| mnda | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ |
| invoice | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 2/3 · ✓ · ✗ |
| resume | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 0/3 · ✗ · ✗ |
| contract-markup | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ |
| contract-finalize | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ |
| eliot-journal | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ | 3/3 · ✓ · ✓ |
| **all 6** | 18/18 · 6/6 · 6/6 | 18/18 · 6/6 · 6/6 | 14/18 · 5/6 · 4/6 |

| Run | Solved | Partial | Rendered | Layout sound | Eff. input | Output | Wall-clock | Tool calls | Breaches/relay |
| :-- | --: | --: | --: | --: | --: | --: | --: | --: | :-- |
| docx-cli r1 | 6/6 | 0 | 6/6 | 6/6 | 904k | 27k | 808 s | 98 | 0/no |
| docx-cli r2 | 6/6 | 0 | 6/6 | 6/6 | 899k | 35k | 943 s | 98 | 0/no |
| docx-cli r3 | 6/6 | 0 | 6/6 | 6/6 | 902k | 33k | 990 s | 98 | 0/no |
| anthropic-docx-skill r1 | 6/6 | 0 | 6/6 | 6/6 | 1.14M | 48k | 1,039 s | 125 | 0/no |
| anthropic-docx-skill r2 | 6/6 | 0 | 6/6 | 6/6 | 1.09M | 48k | 1,214 s | 118 | 0/no |
| anthropic-docx-skill r3 | 6/6 | 0 | 6/6 | 6/6 | 1.22M | 66k | 1,511 s | 128 | 0/no |
| officecli r1 | 4/6 | 2 | 6/6 | 5/6 | 2.00M | 66k | 1,756 s | 174 | 0/no |
| officecli r2 | 5/6 | 1 | 6/6 | 5/6 | 1.78M | 62k | 1,659 s | 157 | 0/no |
| officecli r3 | 5/6 | 1 | 6/6 | 5/6 | 1.99M | 59k | 2,085 s | 165 | 0/no |

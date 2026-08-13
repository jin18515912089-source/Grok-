===== BEGIN T21 =====
STATUS: COMPLETE_TEXT_PROXY
ORIGINAL_DELIVERABLE: 中文问题清单 + 建议英文修订条款（非正式法律意见）
PROXY_DELIVERABLE: 问题清单与互为一致的英文修订

## ANSWER

**性质：** 采购合同审阅笔记，非正式法律意见，不声称适用于所有司法辖区。立场：我方可能是披露方也可能是接收方，条款须双向平衡。

### 问题清单

**P1 定义单向且过宽（条款1）｜优先级：高**  
- 风险：仅保护 Customer→Vendor 的信息；未标记也算机密，接收方几乎无法识别。我方若向对方披露架构/价格/流程，将得不到对等保护。  
- 建议立场：改为双向 Confidential Information；以书面标记或口头披露后合理期限内书面确认的信息为主，口头未确认的排除或收窄。  
- 可退让：对明显的商业秘密（源代码、安全架构、非公开价格）即使未标记仍保护。

**P2 期限无限 + 仅可向 employees 披露（条款2）｜高**  
- 风险：无限期不适合一般机密；禁止向关联公司与顾问披露，与双方都可能使用关联公司/专业顾问的背景冲突。  
- 建议：一般信息保密 **3 年**；商业秘密在仍构成商业秘密期间继续保护。允许向关联公司、外部律师/会计师及有必要知悉的承包商披露，前提是书面保密义务不低于本协议。  
- 可退让：核心安全架构可 5 年。

**P3 对方残留记忆例外过大（条款3）｜高**  
- 风险：Customer 可使用 unaided memory 中的任意 idea/know-how，几乎掏空我方技术披露。若对等给 Vendor，双方 IP 都会被掏空。  
- 建议：删除或收窄为不包含源代码、详细设计、客户数据、非公开价格。  
- 可退让：允许使用残留的**一般**行业技能，但禁止使用对方文件与详细规格。

**P4 不可挽回损害 + 自动禁令无需证明（条款4）｜中高**  
- 风险：仅 Vendor 承认、仅 Customer 自动获禁令，不平衡；“无需证明实际损害”在部分程序中过分。  
- 建议：改为双方承认违约可能造成不可挽回损害，**可申请**禁令救济，但仍须遵守适用程序法，不排除需作必要证明。  
- 可退让：保留“不可挽回损害”表述，删除“automatically / without proof”。

**P5 管辖单向（条款5）｜中**  
- 风险：Vendor 接受 Customer 主营业地排他法院 + State A 法，我方作为 Vendor 或 Customer 时不对等。背景首选中立地仲裁。  
- 建议：中立地仲裁；规则与地待确认。  
- 可退让：State A 实体法 + 中立仲裁地。

**P6 无例外（条款6）｜高**  
- 风险：已公开、已合法知悉、独立开发、第三方合法取得均不排除，接收方无法运营。  
- 建议：恢复标准四项排除，并加合法强制披露（事先通知，法律允许时）。  
- 可退让：强制披露通知在法律禁止时可免除。

### 建议英文修订（相互一致、保持双向）

1. “Confidential Information” means non-public information disclosed by either party (a “Disclosing Party”) to the other (a “Receiving Party”) that is marked confidential or, if disclosed orally, summarized in writing within 15 days; provided that source code, security architecture, non-public pricing, and customer processes are Confidential Information whether or not marked.

2. The Receiving Party shall protect Confidential Information using at least reasonable care for three (3) years from disclosure; trade secrets shall be protected for so long as they remain trade secrets under applicable law. Disclosure is permitted to employees, affiliates, and professional advisers with a need to know who are bound by written confidentiality obligations no less protective than this Agreement.

3. Except for general skills and knowledge retained in unaided memory that do not include the Disclosing Party’s source code, detailed designs, customer data, or non-public pricing, neither party may use the other party’s Confidential Information other than for the integration evaluation. Clause 3 of the draft is deleted to that extent.

4. Each party acknowledges that a breach of this Agreement may cause irreparable harm. Either party may seek injunctive relief in accordance with applicable procedural law. Nothing requires a court or tribunal to grant relief without such proof as that forum requires. Automatic entitlement language is removed.

5. This Agreement is governed by the laws of \[TO CONFIRM: State A / other\], without regard to conflicts principles. Disputes shall be finally resolved by arbitration in a mutually agreed neutral seat \[TO CONFIRM\], in English. The exclusive-court submission in the draft is deleted.

6. Confidential Information does not include information that the Receiving Party can demonstrate: (a) is or becomes public other than by breach; (b) was already known without duty of confidentiality; (c) is independently developed without use of the Disclosing Party’s information; or (d) is lawfully received from a third party without duty of confidentiality. Compelled disclosure is permitted to the extent legally required, with prior notice where legally allowed.

## ASSUMPTIONS_AND_UNKNOWNS
仲裁机构、中立地、State A 是否可接受均为待确认。非正式法律意见。

## UNMET_ARTIFACT_REQUIREMENTS
无。

## SELF_CHECK
- 每项含风险、优先级、立场、退让。
- 修订全程 either party / Receiving Party，未改成单向 Vendor 义务。

===== END T21 =====

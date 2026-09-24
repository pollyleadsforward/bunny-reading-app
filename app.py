
import io
import json
import re
from pathlib import Path
from html import escape

import requests
import streamlit as st
import streamlit.components.v1 as components
from bs4 import BeautifulSoup
from docx import Document
from pypdf import PdfReader
from supabase import create_client


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Bunny Reading",
    page_icon="🐰",
    layout="centered",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
ASSET_DIR = BASE_DIR / "assets"


# =========================================================
# CI / THEME CONSTANTS
# =========================================================
CI = {
    "pearl": "#F7F1EA",
    "rose": "#E8B7C8",
    "lavender": "#CFC7E8",
    "blue": "#BFD7EA",
    "sage": "#C9D6C1",
    "gold": "#DCC8A1",
    "taupe": "#75655D",
    "brown": "#8A7468",
    "muted": "#9A8B82",
    "white": "#FFFFFF",
}


# =========================================================
# BOOK #1 DATA
# Source: Head of Personal Loan Product Management Interview Review
# Keep book content separate from UI logic.
# =========================================================
DEMO_BOOKS = [{'book_id': 'technical-product-manager-ai-llm',
  'title': 'Technical Product Manager + AI/LLM',
  'subtitle': '30-Day Intensive Bootcamp · Full Book',
  'author': 'Personal Study Library',
  'content_type': 'Book',
  'category': 'Technical Product Management',
  'description': '30-day TPM bootcamp covering software/system foundations, technical delivery, '
                 'AI/ML/LLM, platform product, production readiness, Chinese technical communication, '
                 'and interview readiness.',
  'cover_emoji': '📗',
  'chapters': [{'chapter_id': 'tpm-day-1',
                'chapter_title': 'Day 1 — บทบาท Technical Product Manager',
                'order': 1,
                'content': 'Learning Objective\n'
                           ' เข้าใจว่าตำแหน่งนี้สร้าง value ตรงไหน และแตกต่างจาก PM, PO, BA, Project '
                           'Manager และ Engineer อย่างไร\n'
                           '\n'
                           '1.1 แก่นของงาน\n'
                           'Technical Product Manager (TPM) เป็นเจ้าของ “ปัญหา + ผลลัพธ์ + '
                           'การตัดสินใจเชิงเทคนิคในระดับ Product” โดย\n'
                           'ไม่จำเป็นต้องเป็นคนลงมือ implement เอง สิ่งสำคัญคือสามารถเชื่อม Business '
                           'Goal กับ Architecture, Data, API,\n'
                           'Security, Reliability และ Delivery ได้\n'
                           '\n'
                           '1.2 เส้นแบ่งของบทบาท\n'
                           'PM เน้น Why/What, Engineer เน้น How, TPM ต้องเข้าใจ Why/What และเข้าใจ How '
                           'มากพอที่จะประเมิน trade-off\n'
                           'และถามคำถามถูกจุด ส่วน Project Manager เน้น timeline/resource/dependency '
                           'และ BA เน้น\n'
                           'process/requirement detail แม้ในองค์กรจริงบทบาทจะ overlap กันได้\n'
                           '\n'
                           '1.3 สิ่งที่ TPM ต้องถือในหัวพร้อมกัน\n'
                           'User value, business outcome, technical feasibility, delivery risk, '
                           'security/compliance, performance,\n'
                           'scalability, cost และ operability หลัง go-live '
                           'ทั้งหมดต้องถูกพิจารณาเป็นระบบเดียว ไม่ใช่แยกเป็นฝ่าย ๆ\n'
                           '\n'
                           '1.4 ตัวอย่างจากคำขอ Business\n'
                           'Business บอก “อยากให้ลูกค้าเห็นสถานะสินเชื่อ real-time” TPM ต้องถามต่อ: '
                           'source of truth อยู่ไหน? real-time\n'
                           'หมายถึงกี่วินาที? update event มาจากระบบใด? ถ้า downstream ล่มจะ fallback '
                           'อย่างไร? มี SLA หรือไม่? มี PII อะไร\n'
                           'ถูกส่งผ่าน? metric หลัง launch คืออะไร?\n'
                           '\n'
                           '1.5 Mental model\n'
                           'ทุก feature ให้คิดเป็น 7 ชั้น: User -> Interface -> Service/API -> Logic -> '
                           'Data -> Dependencies -> Operations. ถ้า\n'
                           'คุณวาด 7 ชั้นนี้ได้ คุณเริ่มคิดแบบ Technical Product แล้ว\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Requirement นี้คืออะไร | 这个需求是什么？\n'
                           '\n'
                           'Business ต้องการแก้ปัญหาอะไร | 业务部门想解决什么问题？\n'
                           '\n'
                           'ขอบเขตของโปรเจกต์นี้คืออะไร | 这个项目的范围是什么？\n'
                           '\n'
                           'เราต้องยืนยัน requirement ก่อน | 我们需要先确认需求。\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'เป้าหมายของ product นี้คืออะไร | 这个产品的目标是什么？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Requirement 需求\n'
                           '\uf0b7 | Product 产品\n'
                           '\uf0b7 | Project 项目\n'
                           '\uf0b7 | Scope 范围\n'
                           '\uf0b7 | Goal 目标\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. อธิบาย TPM ใน 3 ประโยคโดยไม่ใช้คำว่า “คนประสานงาน”\n'
                           '\n'
                           '2. ถ้า Business ขอ chatbot คุณจะถามอะไร 5 ข้อก่อนเริ่มทำ\n'
                           '\n'
                           '3. วาด 7 layers ของ feature ที่คุณคุ้นเคยหนึ่ง feature\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-2',
                'chapter_title': 'Day 2 — Frontend / Backend / Client-Server',
                'order': 2,
                'content': 'Learning Objective\n'
                           ' เข้าใจการไหลของ request ตั้งแต่หน้าจอผู้ใช้ไปถึง backend และ data source\n'
                           '\n'
                           '2.1 Frontend คืออะไร\n'
                           'Frontend คือส่วนที่ user สัมผัส เช่น mobile app, web app, admin portal '
                           'หน้าที่หลักคือรับ input แสดงผล จัดการ\n'
                           'state ฝั่ ง client และเรียก backend ผ่าน network ไม่ควรถือ business logic '
                           'สำคัญทั้งหมดไว้ที่ client เพราะ\n'
                           'แก้ไข/ควบคุม/security ยากกว่า\n'
                           '\n'
                           '2.2 Backend คืออะไร\n'
                           'Backend คือ service ฝั่ ง server ที่รับ request ตรวจสอบสิทธิ์ ประมวลผล '
                           'business logic อ่าน/เขียน database เรียก\n'
                           'downstream system และส่ง response กลับ Frontend หนึ่ง product อาจมี backend '
                           'หลาย service\n'
                           '\n'
                           '2.3 Client-Server flow\n'
                           'ตัวอย่าง: User กด “ดูวงเงิน” -> Frontend ตรวจ input -> ส่ง HTTPS request -> '
                           'API Gateway -> Backend ->\n'
                           'Database/Core system -> Backend แปลงผล -> Frontend render ให้ user เห็น\n'
                           '\n'
                           '2.4 State และ Session\n'
                           'บางระบบเป็น stateless: ทุก request มีข้อมูลพอให้ server ทำงานได้เอง '
                           'บางระบบใช้ session/token เพื่อจำว่า user\n'
                           'login แล้ว TPM ต้องรู้ว่า session หมดอายุเมื่อไร, user experience เมื่อ '
                           'token expire เป็นอย่างไร และมี security\n'
                           'implication อะไร\n'
                           '\n'
                           '2.5 Debug mindset\n'
                           'เมื่อหน้าจอผิด อย่าพูดว่า “ระบบพัง” ให้แยก: UI render ผิด? request ไม่ออก? '
                           'API error? backend logic ผิด? data\n'
                           'source ผิด? network timeout? การแยก layer ทำให้ incident triage เร็วขึ้น\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           ' เป็นปัญหาของ Frontend หรือ\n'
                           '                                       是前端的问题还是后端的问题？\n'
                           ' Backend\n'
                           '\n'
                           'Frontend แสดงข้อมูลไม่ถูกต้อง | 前端显示的数据不正确。\n'
                           '\n'
                           'Backend ส่งข้อมูลอะไรกกลับมา | 后端返回了什么数据？\n'
                           '\n'
                           'ปัญหาเกิดที่ชั้นไหนของระบบ | 问题发生在哪一层？\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ผู้ใช้กดปุ่มแล้วไม่มี response | 用户点击按钮后没有响应。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Frontend 前端\n'
                           '\uf0b7 | Backend 后端\n'
                           '\uf0b7 | System 系统\n'
                           '\uf0b7 | Response 响应\n'
                           '\uf0b7 | User 用户\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. อธิบาย flow ตอน user login ให้ครบอย่างน้อย 5 steps\n'
                           '\n'
                           '2. ถ้า UI แสดงยอดเงินเก่า คุณจะตรวจ layer ไหนบ้าง\n'
                           '\n'
                           '3. Frontend ควรเก็บ password plain text หรือไม่ เพราะอะไร\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-3',
                'chapter_title': 'Day 3 — API / REST / JSON / Authentication',
                'order': 3,
                'content': 'Learning Objective\n'
                           ' อ่าน API concept และคุยกับ developer เรื่อง request/response, error และ '
                           'authentication ได้\n'
                           '\n'
                           '3.1 API คือสัญญาการสื่อสาร\n'
                           'API คือ interface/contract '
                           'ที่กำหนดว่าระบบหนึ่งจะขอให้อีกระบบทำอะไรได้อย่างไร Contract ประกอบด้วย '
                           'endpoint,\n'
                           'method, parameters, headers, body, response schema, error code และ '
                           'authentication\n'
                           '\n'
                           '3.2 REST mental model\n'
                           'REST API มักใช้ HTTP methods: GET อ่านข้อมูล, POST สร้าง/สั่ง action, PUT '
                           'แทนที่ resource, PATCH แก้บางส่วน,\n'
                           'DELETE ลบ โดย URI ควรสื่อถึง resource เช่น /customers/123/loans\n'
                           '\n'
                           '3.3 JSON\n'
                           'JSON เป็นรูปแบบ key-value ที่ใช้แลกข้อมูล ตัวอย่าง response มี '
                           'applicationId, status, amount. TPM ควรอ่าน\n'
                           'nested JSON, optional field, null และ data type ได้ เพื่อ review contract '
                           'และ acceptance criteria\n'
                           '\n'
                           ' GET /api/v1/loans/12345\n'
                           ' Authorization: Bearer <token>\n'
                           '\n'
                           ' HTTP/1.1 200 OK\n'
                           ' {\n'
                           '   "loanId": "12345",\n'
                           '   "status": "APPROVED",\n'
                           '   "amount": 500000\n'
                           ' }\n'
                           '\n'
                           '3.4 HTTP status\n'
                           '2xx สำเร็จ, 4xx ปัญหาจาก request/client/auth, 5xx ปัญหาฝั่ ง '
                           'server/downstream แต่ต้องดู error body ประกอบ\n'
                           'อย่าใช้ status code อย่างเดียวเป็น root cause\n'
                           '\n'
                           '3.5 Authentication vs Authorization\n'
                           'Authentication = คุณคือใคร เช่น OAuth/OIDC/login/token; Authorization = '
                           'คุณมีสิทธิ์ทำอะไร เช่น\n'
                           'role/permission/scope. Token เช่น JWT อาจมี claims ระบุ user/role/expiry '
                           'แต่ TPM ไม่ต้องเขียน token เอง ต้อง\n'
                           'เข้าใจ lifecycle และ risk\n'
                           '\n'
                           '3.6 Idempotency และ Retry\n'
                           'การ retry POST payment ซ้ำอาจเกิด double charge จึงต้องมี idempotency key '
                           'หรือ design ที่ป้องกัน duplicate.\n'
                           'TPM ควรถามทุก transaction สำคัญว่า retry แล้วเกิดผลซ้ำหรือไม่\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'API นี้เรียกใช้งานอย่างไร | 这个 API 是怎么调用的？\n'
                           '\n'
                           'Request ต้องส่ง parameter อะไรบ้าง | 这个请求需要传哪些参数？\n'
                           '\n'
                           'API ส่ง response อะไรกลับมา | 这个 API 返回什么数据？\n'
                           '\n'
                           'ต้องใช้ token แบบไหน | 需要使用什么令牌？\n'
                           '\n'
                           'กรณี error ระบบส่ง error code อะไร | 发生错误时，系统返回什么错误码？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | API/Interface 接口\n'
                           '\uf0b7 | Request 请求\n'
                           '\uf0b7 | Parameter 参数\n'
                           '\uf0b7 | Token 令牌\n'
                           '\uf0b7 | Error code 错误码\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. GET กับ POST ต่างกันอย่างไรในมุม product\n'
                           '\n'
                           '2. Authentication กับ Authorization ต่างกันอย่างไร\n'
                           '\n'
                           '3. ทำไม retry payment ต้องระวัง duplicate\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-4',
                'chapter_title': 'Day 4 — Database / SQL / Data Flow',
                'order': 4,
                'content': 'Learning Objective\n'
                           ' เข้าใจ relational data, key, join, transaction, source of truth และ query '
                           'พื้นฐาน\n'
                           '\n'
                           '4.1 Table และ Key\n'
                           'Relational database เก็บข้อมูลใน table; row คือ record; column คือ '
                           'attribute. Primary Key ระบุ record ไม่ซ้ำ\n'
                           'เช่น customer_id. Foreign Key ใช้เชื่อม relation เช่น loan.customer_id -> '
                           'customer.customer_id\n'
                           '\n'
                           '4.2 SQL ที่ TPM ควรอ่านได้\n'
                           'SELECT เพื่ออ่าน, WHERE กรอง, JOIN รวมข้อมูลจากหลาย table, GROUP BY สรุป, '
                           'ORDER BY เรียง. ไม่จำเป็นต้อง\n'
                           'เป็น DBA แต่ควร query ข้อมูลเพื่อตรวจ incident/metric เบื้องต้นได้\n'
                           '\n'
                           ' SELECT l.loan_id, l.amount, c.name\n'
                           ' FROM loan l\n'
                           ' JOIN customer c ON c.customer_id = l.customer_id\n'
                           " WHERE l.status = 'ACTIVE'\n"
                           ' ORDER BY l.amount DESC;\n'
                           '\n'
                           '4.3 Transaction และ Consistency\n'
                           'ธุรกรรมบางชุดต้องสำเร็จทั้งหมดหรือไม่สำเร็จเลย เช่น debit/credit. แนวคิด '
                           'ACID ช่วยให้เข้าใจ\n'
                           'atomicity/consistency. TPM ต้องถามว่าถ้าขั้นตอนกลาง fail จะ rollback หรือ '
                           'compensate อย่างไร\n'
                           '\n'
                           '4.4 Source of Truth\n'
                           'ข้อมูลเดียวกันอาจ copy อยู่หลายระบบ ต้องกำหนด authoritative source เช่น '
                           'Customer Master เป็น source of\n'
                           'truth ของชื่อ ส่วน Loan Core เป็น source of truth ของ loan status. '
                           'ถ้าไม่ชัดจะเกิด reconciliation issue\n'
                           '\n'
                           '4.5 Eventual consistency\n'
                           'Distributed systems บางครั้งข้อมูลไม่ตรงกันชั่วคราว เช่น update ใน core '
                           'แล้ว cache/search index ตามหลัง 30\n'
                           'วินาที TPM ต้องแปลง technical behavior เป็น UX ที่ยอมรับได้และ SLA ที่ชัด\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ข้อมูลนี้มาจาก database ไหน | 这个数据来自哪个数据库？\n'
                           '\n'
                           'ระบบไหนเป็นแหล่งข้อมูลหลัก | 哪个系统是主要数据源？\n'
                           '\n'
                           'กรุณาตรวจสอบข้อมูลใน database | 请检查数据库里的数据。\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Field นี้สามารถเป็นค่าว่างได้ไหม | 这个字段可以为空吗？\n'
                           '\n'
                           'ข้อมูลสองระบบไม่ตรงกัน | 两个系统的数据不一致。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Database 数据库\n'
                           '\uf0b7 | Data 数据\n'
                           '\uf0b7 | Field 字段\n'
                           '\uf0b7 | Query 查询\n'
                           '\uf0b7 | Consistency 一致性\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. Primary key กับ foreign key คืออะไร\n'
                           '\n'
                           '2. ถ้าชื่อใน CRM กับ Core ไม่ตรง จะถาม source of truth อย่างไร\n'
                           '\n'
                           '3. อธิบาย eventual consistency ด้วยตัวอย่างของตัวเอง\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-5',
                'chapter_title': 'Day 5 — Architecture / Services / Integration',
                'order': 5,
                'content': 'Learning Objective\n'
                           ' อ่าน architecture diagram และระบุ dependency, bottleneck, sync/async '
                           'integration ได้\n'
                           '\n'
                           '5.1 Monolith vs Microservices\n'
                           'Monolith รวมหลาย function ใน application เดียว ทำง่ายช่วงแรกแต่ '
                           'deploy/scale แยกยาก; microservices แยก\n'
                           'ตาม capability deploy/scale แยกได้ แต่เพิ่ม complexity เช่น network, '
                           'observability, data consistency. TPM ไม่\n'
                           'ควรเชื่อว่า microservices “ดีกว่าเสมอ”\n'
                           '\n'
                           '5.2 API Gateway\n'
                           'API Gateway เป็นประตูหน้าของ backend APIs ช่วย routing, authentication, '
                           'rate limiting, logging หรือ policy\n'
                           'enforcement. ถ้า gateway ล่ม service หลังบ้านที่ปกติดีก็เข้าถึงไม่ได้\n'
                           '\n'
                           '5.3 Synchronous vs Asynchronous\n'
                           'Sync: caller รอ response เหมาะกับงานต้องรู้ผลทันที; Async: ส่ง '
                           'message/event แล้ว process ภายหลังผ่าน\n'
                           'queue/broker เหมาะกับงานยาวหรือ decouple systems เช่น ส่ง notification หลัง '
                           'approve\n'
                           '\n'
                           ' Synchronous:\n'
                           ' Frontend -> API -> Loan Service -> Core -> Response\n'
                           '\n'
                           ' Asynchronous:\n'
                           ' Loan Service -> Event: LoanApproved -> Queue -> Notification Service\n'
                           '\n'
                           '5.4 Queue / Event\n'
                           'Message queue ช่วย buffer load และ retry; event-driven architecture ให้ '
                           'system publish event เช่น\n'
                           'LoanApproved แล้วระบบอื่น subscribe. TPM ต้องถาม delivery guarantee, '
                           'duplicate event, ordering และ\n'
                           'dead-letter handling ใน use case สำคัญ\n'
                           '\n'
                           '5.5 Cache\n'
                           'Cache เก็บข้อมูลที่เรียกบ่อยเพื่อเร็วขึ้น ลด load แต่เสี่ยงข้อมูล stale. '
                           'Product decision คือยอม stale ได้กี่วินาที/นาที\n'
                           '\n'
                           '5.6 Bottleneck & SPOF\n'
                           'Single Point of Failure คือ component เดียวล่มแล้วระบบหลักหยุด. Bottleneck '
                           'คือจุดที่จำกัด throughput.\n'
                           'Architecture review ต้องหา dependency critical และ fallback\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           '    ช่วยอธิบาย architecture ให้ฉันฟัง\n'
                           '                                         请给我解释一下系统架构。\n'
                           '    หน่อย\n'
                           '\n'
                           'ข้อมูลไหลจากระบบไหนไปยังระบบไหน 数据从哪个系统传到哪个系统？\n'
                           '\n'
                           'API นี้เชื่อมต่อกับระบบอะไร | 这个 API 连接哪个系统？\n'
                           '\n'
                           '    ขั้นตอนนี้เป็น synchronous หรือ\n'
                           '                                         这个步骤是同步还是异步？\n'
                           '    asynchronous\n'
                           '\n'
                           'ถ้าระบบนี้ล่มมี fallback ไหม | 如果这个系统不可用，有备用方案吗？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Architecture 架构\n'
                           '\uf0b7 | Dependency 依赖\n'
                           '\uf0b7 | Synchronous 同步\n'
                           '\uf0b7 | Asynchronous 异步\n'
                           '\uf0b7 | Cache 缓存\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. ยก use case ที่ควร sync 1 อันและ async 1 อัน\n'
                           '\n'
                           '2. Cache มีข้อดี/ความเสี่ยงอะไร\n'
                           '\n'
                           '3. วาด dependency ของระบบสินเชื่อหนึ่ง flow\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-6',
                'chapter_title': 'Day 6 — Cloud / Environment / Deployment / Git / CI-CD',
                'order': 6,
                'content': 'Learning Objective\n'
                           ' เข้าใจเส้นทาง code จาก developer ไป production และศัพท์ DevOps ที่ TPM '
                           'ต้องคุยได้\n'
                           '\n'
                           '6.1 Environment\n'
                           'DEV ใช้พัฒนา, SIT ทดสอบ integration, UAT ให้ business ยืนยัน, '
                           'Staging/Pre-prod ใกล้ production, Production\n'
                           'ใช้งานจริง. ชื่อ environment ต่างกันตามองค์กร แต่หลักคือแยกความเสี่ยง\n'
                           '\n'
                           '6.2 Cloud primitives\n'
                           'Compute รัน application, storage เก็บไฟล์/object, database เก็บ structured '
                           'data, network เชื่อมระบบ, load\n'
                           'balancer กระจาย traffic, autoscaling เพิ่ม/ลด capacity. TPM ต้องเข้าใจผลต่อ '
                           'cost/reliability มากกว่ารายละเอียด\n'
                           'command\n'
                           '\n'
                           '6.3 Git\n'
                           'Git เก็บ version ของ code. Developer ทำ branch -> commit -> pull '
                           'request/merge request -> review ->\n'
                           'merge. TPM ควรรู้ว่า feature ไหนอยู่ branch/version ไหนเพื่อเชื่อม roadmap '
                           'กับ release\n'
                           '\n'
                           '6.4 CI/CD\n'
                           'Continuous Integration รัน build/test อัตโนมัติเมื่อ code เปลี่ยน; '
                           'Continuous Delivery/Deployment ทำ package\n'
                           'และส่งไป environment อย่าง repeatable ลด manual error. Pipeline fail '
                           'ต้องรู้ว่า stage ไหน fail\n'
                           '\n'
                           '6.5 Deployment strategies\n'
                           'Rolling update เปลี่ยนทีละ instance; Blue-Green สลับ environment; Canary '
                           'ปล่อยให้ user กลุ่มเล็กก่อน. TPM\n'
                           'เลือก strategy ตาม risk, traffic, rollback speed\n'
                           '\n'
                           '6.6 Configuration & Secrets\n'
                           'Config ควรแยกตาม environment เช่น endpoint; secrets เช่น password/API key '
                           'ไม่ควร hard-code ใน source\n'
                           'code. TPM ควรถาม secret management และ access control\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ตอนนี้อยู่ environment ไหน | 现在是哪个环境？\n'
                           '\n'
                           'Version นี้จะขึ้น production เมื่อไร | 这个版本什么时候上线？\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Pipeline fail ที่ขั้นตอนไหน | 流水线在哪个步骤失败了？\n'
                           '\n'
                           'ถ้ามีปัญหาสามารถ rollback ได้ไหม | 如果出现问题，可以回滚吗？\n'
                           '\n'
                           'เราจะปล่อยแบบ canary ก่อน | 我们先进行金丝雀发布。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Environment 环境\n'
                           '\uf0b7 | Deployment 部署\n'
                           '\uf0b7 | Go-live 上线\n'
                           '\uf0b7 | Rollback 回滚\n'
                           '\uf0b7 | Pipeline 流水线\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. CI กับ CD คืออะไร\n'
                           '\n'
                           '2. Canary release เหมาะเมื่อไร\n'
                           '\n'
                           '3. ทำไม secrets ไม่ควรอยู่ใน code\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-7',
                'chapter_title': 'Day 7 — Observability / Incident / Reliability',
                'order': 7,
                'content': 'Learning Objective\n'
                           ' เข้าใจ logs, metrics, traces, SLA/SLO และวิธีจัดการ incident '
                           'อย่างเป็นระบบ\n'
                           '\n'
                           '7.1 Observability 3 pillars\n'
                           'Logs บอกเหตุการณ์ละเอียด, Metrics บอกค่าตัวเลขตามเวลา, Traces ติดตาม '
                           'request ข้ามหลาย services. สามอย่าง\n'
                           'ช่วยตอบว่า “เกิดอะไร ที่ไหน เมื่อไร และกระทบอะไร”\n'
                           '\n'
                           '7.2 Golden signals\n'
                           'Latency, Traffic, Errors, Saturation เป็นสัญญาณหลักที่ใช้ monitor service. '
                           'Product metric และ system metric\n'
                           'ต้องเชื่อมกัน เช่น checkout conversion ลดเพราะ latency เพิ่ม\n'
                           '\n'
                           '7.3 SLA / SLO / SLI\n'
                           'SLI คือสิ่งที่วัด เช่น success rate; SLO คือ target ภายใน เช่น 99.9%; SLA '
                           'คือคำมั่นกับลูกค้า/คู่สัญญาที่อาจมีผลเชิง\n'
                           'พาณิชย์. TPM ต้องระวังอย่าใช้คำสลับกัน\n'
                           '\n'
                           '7.4 Incident process\n'
                           'Detect -> Triage -> Contain -> Communicate -> Diagnose -> Fix -> Validate '
                           '-> Recover -> Postmortem. ระหว่าง\n'
                           'incident ให้แยก temporary workaround กับ permanent fix\n'
                           '\n'
                           '7.5 Root Cause vs Symptom\n'
                           'CPU สูงอาจเป็น symptom ไม่ใช่ root cause. Root cause อาจเป็น query '
                           'ใหม่ที่ไม่มี index. การแก้ที่ symptom อย่าง\n'
                           'restart อาจช่วยชั่วคราวแต่ปัญหากลับมา\n'
                           '\n'
                           '7.6 Error budget\n'
                           'ถ้า SLO 99.9% จะยอม downtime/error ได้บางส่วน Error budget ใช้บาลานซ์ '
                           'feature velocity กับ reliability: ถ้าใช้\n'
                           'budget หมดควรชะลอ risky release\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'กรุณาส่ง error log | 请提供错误日志。\n'
                           '\n'
                           'Root cause คืออะไร | 根本原因是什么？\n'
                           '\n'
                           'มีผลกระทบกับลูกค้ากี่คน | 影响了多少客户？\n'
                           '\n'
                           'Response time เพิ่มขึ้นมาก | 响应时间明显增加了。\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'เราต้องทำ postmortem หลัง incident | 事故后我们需要做复盘。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Log 日志\n'
                           '\uf0b7 | Monitoring 监控\n'
                           '\uf0b7 | Root cause 根本原因\n'
                           '\uf0b7 | Latency 延迟\n'
                           '\uf0b7 | Postmortem 复盘\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. Logs, metrics, traces ต่างกันอย่างไร\n'
                           '\n'
                           '2. SLO 99.9% หมายความว่าอะไรในเชิง product\n'
                           '\n'
                           '3. เขียน incident questions 8 ข้อที่คุณจะถาม engineer\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่\n'
                           '\n'
                           'WEEK 2 - PRODUCT DELIVERY & TECHNICAL\n'
                           '                    EXECUTION\n'
                           '        เปลี่ยน Requirement ให้เป็น Production Delivery'},
               {'chapter_id': 'tpm-day-8',
                'chapter_title': 'Day 8 — Product Discovery & Problem Framing',
                'order': 8,
                'content': 'Learning Objective\n'
                           ' เปลี่ยน feature request ให้เป็นปัญหา/ผลลัพธ์ที่วัดได้\n'
                           '\n'
                           '8.1 Problem before solution\n'
                           'คำขอ “ทำ chatbot” เป็น solution request. ต้องย้อนถาม user job, pain, '
                           'frequency, impact, alternative และ\n'
                           'constraint จนได้ problem statement เช่น “เจ้าหน้าที่ใช้เวลาเฉลี่ย 15 '
                           'นาทีค้น policy และ 12% ของเคสต้องถาม\n'
                           'senior ซ้ำ”\n'
                           '\n'
                           '8.2 User + Job to be Done\n'
                           'ระบุ user segment และ job เช่น Loan Officer ต้องการ “หาคำตอบ policy '
                           'ที่เชื่อถือได้ก่อนตอบลูกค้า” ไม่ใช่เพียง “ใช้\n'
                           'chatbot”\n'
                           '\n'
                           '8.3 Outcome metric\n'
                           'Output = feature shipped; Outcome = behavior/business changed. ตัวอย่าง: '
                           'median search time 15 -> 2 นาที,\n'
                           'first-contact resolution 65% -> 85%, incorrect-answer rate <1%\n'
                           '\n'
                           '8.4 Constraints\n'
                           'Regulation, security, legacy dependency, budget, data quality, latency, '
                           'vendor contract ล้วนเป็น constraint\n'
                           'ที่ต้องใส่ใน discovery ไม่ใช่รอเจอตอน build\n'
                           '\n'
                           '8.5 Assumption mapping\n'
                           'แบ่ง assumption เป็น desirability, feasibility, viability, usability '
                           'แล้วเลือก assumption ที่เสี่ยงสุดมาทดสอบก่อน\n'
                           'ลดการลงทุนใน solution ที่ไม่ตอบโจทย์\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ปัญหาหลักของผู้ใช้คืออะไร | 用户的主要问题是什么？\n'
                           '\n'
                           'เราต้องวัดผลลัพธ์ ไม่ใช่แค่ output | 我们需要衡量结果，而不仅仅是产出。\n'
                           '\n'
                           'ข้อสมมติฐานที่เสี่ยงที่สุดคืออะไร | 风险最大的假设是什么？\n'
                           '\n'
                           'ใครคือผู้ใช้หลัก | 谁是主要用户？\n'
                           '\n'
                           'เราจะพิสูจน์ปัญหานี้อย่างไร | 我们怎么验证这个问题？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | User 用户\n'
                           '\uf0b7 | Problem 问题\n'
                           '\uf0b7 | Outcome 结果\n'
                           '\uf0b7 | Assumption 假设\n'
                           '\uf0b7 | Validate 验证\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. เปลี่ยน “อยากได้ chatbot” เป็น problem statement\n'
                           '\n'
                           '2. ตั้ง outcome metric 3 ตัวสำหรับ knowledge assistant\n'
                           '\n'
                           '3. ยก assumption ที่ต้อง test ก่อน build\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-9',
                'chapter_title': 'Day 9 — PRD / Requirements / NFR',
                'order': 9,
                'content': 'Learning Objective\n'
                           ' เขียน requirement ที่ชัด วัดได้ และครอบคลุม technical constraints\n'
                           '\n'
                           '9.1 PRD skeleton\n'
                           'PRD ที่ดีควรมี Context, Problem, Objective, Users, Use Cases, Scope, Out of '
                           'Scope, Functional Requirements,\n'
                           'NFR, Data, Dependencies, Risks, Metrics, Rollout, Open Questions\n'
                           '\n'
                           '9.2 Functional requirements\n'
                           'บอกว่าระบบทำอะไร เช่น user ค้น policy, filter ตาม product, เปิด citation, '
                           'ส่ง feedback. แต่ละ requirement ควร\n'
                           'trace ไปที่ user need\n'
                           '\n'
                           '9.3 NFR\n'
                           'Performance (P95 latency), availability, scalability, security, privacy, '
                           'accessibility, auditability,\n'
                           'recoverability, compatibility. คำว่า “เร็ว” “เสถียร” '
                           'ใช้ไม่ได้ถ้าไม่มีตัวเลข/เงื่อนไข\n'
                           '\n'
                           '9.4 Edge cases\n'
                           'Null, duplicate, timeout, partial failure, stale data, permission denied, '
                           'document missing, concurrent\n'
                           'update, retry หลัง network disconnect ต้องถูกคิดก่อน testing\n'
                           '\n'
                           '9.5 Decision log\n'
                           'Technical decisions ที่สำคัญควรถูกบันทึก: decision, options, reason, '
                           'trade-off, owner, date. ช่วยลดการย้อน\n'
                           'เถียงและรักษาความรู้ทีม\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Requirement นี้ยังไม่ชัดเจน | 这个需求还不够明确。\n'
                           '\n'
                           ' กรุณาระบุ non-functional\n'
                           '                                         请明确非功能性需求。\n'
                           ' requirements\n'
                           '\n'
                           'Response time ต้องไม่เกินสองวินาที | 响应时间必须不超过两秒。\n'
                           '\n'
                           'กรุณายืนยัน scope | 请确认项目范围。\n'
                           '\n'
                           ' กรณี exception นี้ระบบต้องทำอย่างไร 这种异常情况下系统应该怎么处理？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Requirement 需求\n'
                           '\uf0b7 | NFR 非功能性需求\n'
                           '\uf0b7 | Exception 异常\n'
                           '\uf0b7 | Risk 风险\n'
                           '\uf0b7 | Metric 指标\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. เขียน NFR 5 ตัวสำหรับ mobile banking feature\n'
                           '\n'
                           '2. Edge case ต่างจาก happy path อย่างไร\n'
                           '\n'
                           '3. PRD ควรมี Out of Scope เพราะอะไร\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-10',
                'chapter_title': 'Day 10 — User Story / Acceptance Criteria / API Contract',
                'order': 10,
                'content': 'Learning Objective\n'
                           ' ทำ requirement ให้ testable และเชื่อมกับ implementation contract\n'
                           '\n'
                           '10.1 User Story\n'
                           'รูปแบบ As a [user], I want [capability], so that [outcome] ช่วยย้ำ user '
                           'value แต่ไม่จำเป็นต้องบังคับใช้กับทุก\n'
                           'technical enabler\n'
                           '\n'
                           '10.2 Acceptance Criteria\n'
                           'ควรชัดและ testable เช่น Given-When-Then. ระบุ input, state, action, '
                           'expected output, error behavior. หลีก\n'
                           'เลี่ยงคำ subjective เช่น user-friendly โดยไม่ define\n'
                           '\n'
                           '10.3 API contract as product artifact\n'
                           'TPM ควร review field names, data type, mandatory/optional, enum, '
                           'pagination, error contract, backward\n'
                           'compatibility และ versioning เพื่อป้องกัน frontend/backend mismatch\n'
                           '\n'
                           '10.4 Backward compatibility\n'
                           'การเปลี่ยน field/enum อาจทำ client เก่าพัง ต้องวาง deprecation plan, '
                           'versioning หรือ additive change. “แก้\n'
                           'API เล็กน้อย” อาจกระทบหลาย consumer\n'
                           '\n'
                           '10.5 Definition of Done\n'
                           'Done อาจรวม code merged, tests passed, security check, monitoring '
                           'dashboard, runbook, documentation,\n'
                           'migration completed, business sign-off ไม่ใช่แค่ developer บอกเสร็จ\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Acceptance criteria คืออะไร | 验收标准是什么？\n'
                           '\n'
                           'Field นี้เป็น mandatory หรือ optional | 这个字段是必填还是可选？\n'
                           '\n'
                           ' การเปลี่ยน API นี้ backward\n'
                           '这个 API 变更向后兼容吗？\n'
                           ' compatible ไหม\n'
                           '\n'
                           'กรุณาเพิ่ม error scenario | 请补充错误场景。\n'
                           '\n'
                           'Definition of Done ของงานนี้คืออะไร | 这个任务的完成标准是什么？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Acceptance 验收\n'
                           '\uf0b7 | Mandatory 必填\n'
                           '\uf0b7 | Optional 可选\n'
                           '\uf0b7 | Compatible 兼容\n'
                           '\uf0b7 | Scenario 场景\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. เขียน Given-When-Then สำหรับ login fail\n'
                           '\n'
                           '2. เหตุใด API contract ต้องคิด backward compatibility\n'
                           '\n'
                           '3. นิยาม Done สำหรับ feature ที่มี production impact\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-11',
                'chapter_title': 'Day 11 — Testing Strategy',
                'order': 11,
                'content': 'Learning Objective\n'
                           ' เข้าใจ test layers และออกแบบ risk-based testing\n'
                           '\n'
                           '11.1 Test pyramid\n'
                           'Unit test เร็วและเยอะ ทดสอบ function; integration test ทดสอบ component '
                           'interactions; end-to-end test\n'
                           'ทดสอบ flow จริงแต่ช้า/เปราะ. ทีมที่พึ่ง E2E อย่างเดียว feedback จะช้า\n'
                           '\n'
                           '11.2 SIT / UAT / Regression\n'
                           'SIT เน้นระบบคุยกันถูก; UAT ยืนยัน business requirement; regression '
                           'ป้องกันของเก่าพังจากของใหม่. บางองค์กร\n'
                           'ใช้ชื่อแตกต่าง แต่ intent สำคัญกว่า acronym\n'
                           '\n'
                           '11.3 Contract testing\n'
                           'API consumer/provider สามารถใช้ contract tests เพื่อจับ breaking change '
                           'ก่อน integration environment ลด\n'
                           'dependency testing\n'
                           '\n'
                           '11.4 Performance testing\n'
                           'Load test ปริมาณคาดการณ์, stress test เกิน capacity, soak test รันนานหา '
                           'memory leak. TPM ต้องกำหนด\n'
                           'realistic workload และ acceptance\n'
                           '\n'
                           '11.5 Risk-based testing\n'
                           'จัด priority ตาม impact x likelihood: payment duplication, security bypass, '
                           'data corruption ต้อง test ลึกกว่า\n'
                           'ปัญหาสีปุ่ม\n'
                           '\n'
                           '11.6 Test data\n'
                           'Production-like แต่ต้อง mask/anonymize PII; test data ต้องครอบคลุม '
                           'boundary, null, multilingual,\n'
                           'duplicate และ historical edge cases\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Testing เจอ defect กี่รายการ | 测试发现了多少个缺陷？\n'
                           '\n'
                           'Defect นี้ block go-live หรือไม่ | 这个缺陷会影响上线吗？\n'
                           '\n'
                           'เราต้องทำ regression test | 我们需要做回归测试。\n'
                           '\n'
                           'ผล performance test เป็นอย่างไร | 性能测试结果怎么样？\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Test data มีข้อมูลส่วนบุคคลหรือไม่ | 测试数据里有个人信息吗？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Testing 测试\n'
                           '\uf0b7 | Defect 缺陷\n'
                           '\uf0b7 | Regression 回归\n'
                           '\uf0b7 | Performance 性能\n'
                           '\uf0b7 | Test data 测试数据\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. Unit/integration/E2E ต่างกันอย่างไร\n'
                           '\n'
                           '2. ยก critical scenario ที่ต้อง regression\n'
                           '\n'
                           '3. P95 latency ใช้ใน performance acceptance ได้อย่างไร\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-12',
                'chapter_title': 'Day 12 — Estimation / Dependency / Trade-off / Technical Debt',
                'order': 12,
                'content': 'Learning Objective\n'
                           '  ตัดสินใจ scope/time/quality โดยเข้าใจความเสี่ยงเชิงเทคนิค\n'
                           '\n'
                           '12.1 Estimation is uncertainty\n'
                           'Estimate ไม่ใช่สัญญาแม่น 100%; งานใหม่/legacy integration มี uncertainty '
                           'สูง. TPM ควรถาม range,\n'
                           'assumptions และ confidence แทนกดให้ทีมตอบเลขเดียว\n'
                           '\n'
                           '12.2 Dependency map\n'
                           'Internal dependency เช่น data team/API platform; external เช่น '
                           'vendor/regulator. ระบุ owner, need-by date,\n'
                           'failure impact และ contingency เพื่อหา critical path\n'
                           '\n'
                           '12.3 MVP\n'
                           'MVP ไม่ใช่ “ทำของแย่” แต่เป็น minimum scope ที่ทดสอบ value/ลด risk ได้จริง. '
                           'ต้องรักษา security/reliability ขั้น\n'
                           'ต่ำ\n'
                           '\n'
                           '12.4 Technical debt\n'
                           'Shortcut ที่ช่วยเร็ววันนี้แต่เพิ่ม cost/risk ในอนาคต เช่น hard-code, no '
                           'tests, duplicated logic. บาง debt ยอมรับได้\n'
                           'ถ้าบันทึกและมี payback plan\n'
                           '\n'
                           '12.5 Trade-off framework\n'
                           'Quality, latency, cost, time, scope, flexibility มักแลกกัน เช่น model ใหญ่ '
                           'quality สูงแต่แพง/ช้า. TPM ต้องทำ\n'
                           'trade-off visible และเชื่อมกับ product objective\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'งานนี้มี dependency อะไรบ้าง | 这个任务有哪些依赖？\n'
                           '\n'
                           'Development ใช้เวลาประมาณเท่าไร | 开发大概需要多长时间？\n'
                           '\n'
                           'ความเสี่ยงหลักคืออะไร | 主要风险是什么？\n'
                           '\n'
                           'เราสามารถลด scope ได้ไหม | 我们可以缩小范围吗？\n'
                           '\n'
                           ' นี่เป็น technical debt ที่ยอมรับได้หรือ\n'
                           '                                            这个技术债可以接受吗？\n'
                           ' ไม่\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Dependency 依赖\n'
                           '\uf0b7 | Development 开发\n'
                           '\uf0b7 | Risk 风险\n'
                           '\uf0b7 | Technical debt 技术债\n'
                           '\uf0b7 | Delivery 交付\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. ทำ dependency map อย่างน้อย 5 nodes\n'
                           '\n'
                           '2. MVP ต่างจาก prototype อย่างไร\n'
                           '\n'
                           '3. ยก technical debt 2 ตัวและผลระยะยาว\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-13',
                'chapter_title': 'Day 13 — Security / Privacy / Access Control',
                'order': 13,
                'content': 'Learning Objective\n'
                           ' มี security mindset ระดับ Product โดยเฉพาะระบบการเงินและ AI\n'
                           '\n'
                           '13.1 CIA triad\n'
                           'Confidentiality: คนไม่มีสิทธิ์ห้ามเห็น, Integrity: ข้อมูลห้ามถูกแก้ผิด, '
                           'Availability: ระบบต้องพร้อมใช้. Product\n'
                           'requirement มักเกี่ยวทั้งสามด้าน\n'
                           '\n'
                           '13.2 Least privilege\n'
                           'ให้ user/service มีสิทธิ์เท่าที่จำเป็นเท่านั้น ใช้ RBAC/ABAC ตามบริบท. '
                           'Admin access ต้อง audit ได้และแยกหน้าที่เมื่อ\n'
                           'จำเป็น\n'
                           '\n'
                           '13.3 Encryption\n'
                           'Data in transit ใช้ TLS/HTTPS; data at rest ควร encrypt ตาม sensitivity. '
                           'TPM ต้องถาม key management และ\n'
                           'data classification ไม่ใช่แค่ “encrypt แล้ว”\n'
                           '\n'
                           '13.4 PII & Data minimization\n'
                           'เก็บ/ส่งเฉพาะข้อมูลที่จำเป็น ลด exposure. Mask/tokenize เมื่อเป็นไปได้ '
                           'กำหนด retention และ deletion. AI use case\n'
                           'ต้องรู้ว่าข้อมูลถูกส่งไป provider ใดและใช้ train ต่อหรือไม่ตาม '
                           'contract/config\n'
                           '\n'
                           '13.5 Common threats\n'
                           'Injection, broken access control, credential leakage, insecure direct '
                           'object reference, rate abuse, prompt\n'
                           'injection (AI). TPM ควร ensure threat modeling/security review สำหรับ '
                           'critical flow\n'
                           '\n'
                           '13.6 Auditability\n'
                           'ใครทำอะไร เมื่อไร จากระบบไหน ผลอะไร ต้อง trace ได้ โดยเฉพาะ approval, '
                           'model-assisted decision และ privileged\n'
                           'action\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ข้อมูลนี้มีข้อมูลส่วนบุคคลหรือไม่ | 这些数据包含个人信息吗？\n'
                           '\n'
                           'ใครมีสิทธิ์เข้าถึงข้อมูลนี้ | 谁有权限访问这些数据？\n'
                           '\n'
                           'ข้อมูลถูกเข้ารหัสหรือไม่ | 数据是否已经加密？\n'
                           '\n'
                           'เราต้องบันทึก audit log | 我们需要记录审计日志。\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           '    กรุณาตรวจสอบสิทธิ์ของ service\n'
                           '                                         请检查服务账号的权限。\n'
                           '    account\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Permission 权限\n'
                           '\uf0b7 | Encryption 加密\n'
                           '\uf0b7 | Audit 审计\n'
                           '\uf0b7 | Sensitive data 敏感数据\n'
                           '\uf0b7 | Security 安全\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. Authentication/authorization ต่างกันอย่างไรใน security context\n'
                           '\n'
                           '2. Data minimization คืออะไร\n'
                           '\n'
                           '3. AI product ต้องถาม data provider เรื่องใดบ้าง\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-14',
                'chapter_title': 'Day 14 — Release / Migration / Production Readiness',
                'order': 14,
                'content': 'Learning Objective\n'
                           ' วาง go-live ที่ลดความเสี่ยงและมี rollback/monitoring ครบ\n'
                           '\n'
                           '14.1 Release plan\n'
                           'ระบุ scope/version, deployment order, dependencies, window, owners, '
                           'communication, validation, rollback\n'
                           'criteria และ monitoring. Feature flag ช่วยแยก deploy จาก release ได้\n'
                           '\n'
                           '14.2 Data migration\n'
                           'ต้องคิด mapping, transformation, validation, reconciliation, cutover, delta '
                           'data, rollback/forward-fix,\n'
                           'archive. จำนวน record ตรงไม่พอ ต้องตรวจ business-level integrity\n'
                           '\n'
                           '14.3 Production readiness review\n'
                           'Checklist: capacity, security, monitoring, alerts, dashboards, runbook, '
                           'on-call, backup/restore, DR, vendor\n'
                           'support, SLA, support process, known issues\n'
                           '\n'
                           '14.4 Smoke test\n'
                           'หลัง deployment ทดสอบ critical path สั้น ๆ '
                           'เพื่อยืนยันระบบพื้นฐานทำงานก่อนเปิดเต็ม traffic\n'
                           '\n'
                           '14.5 Rollback vs Forward fix\n'
                           'Rollback เหมาะเมื่อย้อน version ได้ปลอดภัย; schema/data migration บางครั้ง '
                           'rollback ยาก ต้อง forward fix. จึง\n'
                           'ต้องคิดก่อน deploy\n'
                           '\n'
                           '14.6 Hypercare\n'
                           'ช่วงหลัง go-live เพิ่ม monitoring/support cadence และ decision threshold '
                           'เพื่อจับ issue เร็ว\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'กรุณายืนยัน deployment plan | 请确认部署计划。\n'
                           '\n'
                           'เราต้องตรวจ reconciliation หลัง\n'
                           '                                      迁移后我们需要做数据核对。\n'
                           'migration\n'
                           '\n'
                           'Rollback criteria คืออะไร | 回滚标准是什么？\n'
                           '\n'
                           'หลัง go-live จะ monitor อะไรบ้าง | 上线后我们要监控哪些指标？\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'คืนนี้มี production deployment | 今晚有生产部署。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Migration 迁移\n'
                           '\uf0b7 | Reconciliation 核对\n'
                           '\uf0b7 | Deployment 部署\n'
                           '\uf0b7 | Rollback 回滚\n'
                           '\uf0b7 | Monitoring 监控\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. เขียน production readiness checklist 10 ข้อ\n'
                           '\n'
                           '2. ทำไม migration rollback ยาก\n'
                           '\n'
                           '3. feature flag ช่วยลด release risk อย่างไร\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่\n'
                           '\n'
                           'WEEK 3 - AI / ML / LLM FOUNDATIONS\n'
                           '      เข้าใจ AI/LLM แบบ Product ที่คุยกับ Engineer รู้เรื่อง'},
               {'chapter_id': 'tpm-day-15',
                'chapter_title': 'Day 15 — AI / ML / Deep Learning / LLM',
                'order': 15,
                'content': 'Learning Objective\n'
                           ' วางแผนที่ของ AI ให้ชัดและรู้ว่า use case ไหนต้องใช้ model จริง\n'
                           '\n'
                           '15.1 AI hierarchy\n'
                           'AI เป็น umbrella; Machine Learning เรียน pattern จาก data; Deep Learning '
                           'ใช้ neural networks หลายชั้น;\n'
                           'Generative AI สร้าง content; LLM คือ model ภาษา large-scale ที่ทำนาย token '
                           'ถัดไปและเรียน representation\n'
                           'จากข้อมูลจำนวนมาก\n'
                           '\n'
                           '15.2 Predictive vs Generative\n'
                           'Predictive ML เช่น default risk score/classification; Generative AI '
                           'เช่นสรุปเอกสาร/ตอบคำถาม. KPI และ risk\n'
                           'ต่างกัน: predictive เน้น precision/recall/AUC; generative เน้น '
                           'correctness/groundedness/relevance/safety\n'
                           '\n'
                           '15.3 Model vs Application\n'
                           'LLM เป็น component ไม่ใช่ product ทั้งหมด. Application ยังมี prompt, '
                           'retrieval, tools, policy, UI, auth,\n'
                           'logging, feedback, monitoring. TPM ต้องออกแบบ system ไม่ใช่เลือก model '
                           'อย่างเดียว\n'
                           '\n'
                           '15.4 When not to use AI\n'
                           'ถ้ากฎชัด deterministic และ risk สูง เช่นคำนวณดอกเบี้ยตามสูตร อาจใช้ '
                           'rule/code ดีกว่า LLM. ใช้ AI เมื่องานมี\n'
                           'ambiguity/unstructured language/pattern ที่ rule ยากและมี '
                           'evaluation/control เพียงพอ\n'
                           '\n'
                           '15.5 AI product risk\n'
                           'Probabilistic output หมายถึงคำตอบเดียวกันอาจแตกต่างและไม่รับประกันถูก 100%; '
                           'Product ต้องมี tolerance,\n'
                           'fallback และ human review ตาม impact\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'นี่เป็น use case ของ machine | 这是机器学习还是大语言模型的应用场\n'
                           'learning หรือ LLM | 景？\n'
                           '\n'
                           'เราไม่จำเป็นต้องใช้ AI กับทุกปัญหา | 不是所有问题都需要使用人工智能。\n'
                           '\n'
                           ' Model เป็นเพียงส่วนหนึ่งของ\n'
                           '                                       模型只是产品的一部分。\n'
                           ' product\n'
                           '\n'
                           'ผลลัพธ์ของ model มีความไม่แน่นอน | 模型输出存在不确定性。\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'เราต้องกำหนด risk tolerance | 我们需要定义风险容忍度。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | AI 人工智能\n'
                           '\uf0b7 | Model 模型\n'
                           '\uf0b7 | Machine learning 机器学习\n'
                           '\uf0b7 | Output 输出\n'
                           '\uf0b7 | Uncertainty 不确定性\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. ยก use case ที่ไม่ควรใช้ LLM 2 ตัว\n'
                           '\n'
                           '2. Model ต่างจาก AI application อย่างไร\n'
                           '\n'
                           '3. Predictive AI กับ Generative AI ใช้ metric ต่างกันอย่างไร\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-16',
                'chapter_title': 'Day 16 — Training / Inference / Transformer Mental Model',
                'order': 16,
                'content': 'Learning Objective\n'
                           ' เข้าใจ lifecycle ของ model และ transformer แบบไม่ลงคณิตศาสตร์เกินจำเป็น\n'
                           '\n'
                           '16.1 Training\n'
                           'Training ปรับ model parameters จากข้อมูลเพื่อเรียน pattern. Foundation '
                           'model ใช้ compute/data มหาศาล;\n'
                           'enterprise product ส่วนใหญ่ไม่ train foundation model เอง แต่ใช้ '
                           'hosted/open model แล้วต่อยอด\n'
                           '\n'
                           '16.2 Inference\n'
                           'Inference คือการใช้ model ที่ train แล้วตอบ input จริง. ใน production TPM '
                           'สนใจ latency, throughput, tokens,\n'
                           'cost, failure rate, quota และ scaling\n'
                           '\n'
                           '16.3 Transformer mental model\n'
                           'Transformer ใช้ attention เพื่อให้อินพุตแต่ละ token “มอง” token '
                           'อื่นที่เกี่ยวข้อง จึงจับ context ได้ดี. ไม่ต้องคำนวณ\n'
                           'matrix แต่ควรรู้ว่า context length และ tokenization มีผลต่อ output/cost\n'
                           '\n'
                           '16.4 Pretraining / Instruction tuning / Alignment\n'
                           'Pretraining เรียนภาษากว้าง ๆ; instruction tuning ช่วยทำตามคำสั่ง; alignment '
                           'techniques ทำให้ตอบสอดคล้อง\n'
                           'preference/safety มากขึ้น. Product teamมักบริโภค model '
                           'ที่ผ่านขั้นเหล่านี้แล้ว\n'
                           '\n'
                           '16.5 Temperature & determinism\n'
                           'Temperature สูงเพิ่มความหลากหลาย; ต่ำทำให้ deterministic มากขึ้น '
                           'แต่ไม่ได้ทำให้ factual โดยอัตโนมัติ. Use case\n'
                           'ธนาคารมักต้องลด randomness และพึ่ง grounding/evaluation\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Model นี้ใช้ข้อมูลอะไรในการ training | 这个模型使用什么数据进行训练？\n'
                           '\n'
                           'Inference ใช้เวลานานเท่าไร | 推理需要多长时间？\n'
                           '\n'
                           'Latency ตอนนี้เท่าไร | 目前延迟是多少？\n'
                           '\n'
                           'เราต้องรองรับกี่ request ต่อวินาที | 我们需要支持每秒多少个请求？\n'
                           '\n'
                           'Temperature ตั้งไว้เท่าไร | 温度参数设置为多少？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Training 训练\n'
                           '\uf0b7 | Inference 推理\n'
                           '\uf0b7 | Latency 延迟\n'
                           '\uf0b7 | Throughput 吞吐量\n'
                           '\uf0b7 | Temperature 温度参数\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. Training กับ inference ต่างกันอย่างไร\n'
                           '\n'
                           '2. Attention ช่วย transformer อย่างไรในภาษาง่าย ๆ\n'
                           '\n'
                           '3. Temperature ต่ำแก้ hallucination ได้ทั้งหมดหรือไม่\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-17',
                'chapter_title': 'Day 17 — Tokens / Context / Prompting',
                'order': 17,
                'content': 'Learning Objective\n'
                           ' เข้าใจข้อจำกัดของ LLM input/output และออกแบบ prompt อย่างเป็นระบบ\n'
                           '\n'
                           '17.1 Tokens\n'
                           'LLM ประมวลผล text เป็น tokens ไม่ใช่คำตรง ๆ ภาษาไทย/จีนอาจ tokenize '
                           'ต่างจากอังกฤษ จำนวน token กระทบ\n'
                           'context limit, latency และ cost\n'
                           '\n'
                           '17.2 Context window\n'
                           'รวม system prompt + user prompt + retrieved docs + conversation + output '
                           'ต้องอยู่ใน context budget. ใส่\n'
                           'ข้อมูลมากเกินไม่ใช่ดีเสมอ เพราะ noise เพิ่มและแพง\n'
                           '\n'
                           '17.3 Prompt structure\n'
                           'Prompt ที่ดีมี Role/Goal, Context, Rules, Input, Output format, Examples '
                           'และ refusal/fallback behavior ตาม\n'
                           'use case. Structured output เช่น JSON ช่วย downstream parsing\n'
                           '\n'
                           '17.4 Prompt injection\n'
                           'ข้อความในเอกสารหรือ user อาจพยายามเปลี่ยน instruction เช่น “ignore previous '
                           'rules”. Product ต้องแยก\n'
                           'trusted instruction, sanitize/tool permissions และไม่ให้ model '
                           'มีสิทธิ์เกินจำเป็น\n'
                           '\n'
                           '17.5 Prompt versioning\n'
                           'Prompt คือ production artifact ต้อง version, test และ monitor เช่นเดียวกับ '
                           'code เพราะแก้ prompt 1 บรรทัดอาจ\n'
                           'เปลี่ยน behavior\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'คำถามนี้ใช้กี่ token | 这个问题使用了多少个 token？\n'
                           '\n'
                           'Context ยาวเกินไป | 上下文太长了。\n'
                           '\n'
                           'กรุณาใช้รูปแบบ JSON ใน response | 请用 JSON 格式返回结果。\n'
                           '\n'
                           'เราต้องป้องกัน prompt injection | 我们需要防止提示词注入。\n'
                           '\n'
                           'Prompt นี้เป็น version ไหน | 这个提示词是哪个版本？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Context 上下文\n'
                           '\uf0b7 | Prompt 提示词\n'
                           '\uf0b7 | Injection 注入\n'
                           '\uf0b7 | Format 格式\n'
                           '\uf0b7 | Version 版本\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. อะไรอยู่ใน context budget บ้าง\n'
                           '\n'
                           '2. Prompt ที่ดีควรมีองค์ประกอบอะไร\n'
                           '\n'
                           '3. ทำไม prompt ต้อง version และ regression test\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-18',
                'chapter_title': 'Day 18 — Embeddings / Vector Search',
                'order': 18,
                'content': 'Learning Objective\n'
                           ' เข้าใจ semantic retrieval ที่เป็นพื้นฐาน RAG\n'
                           '\n'
                           '18.1 Embedding\n'
                           'Embedding แปลงข้อความ/วัตถุเป็น vector ใน space '
                           'ที่ความหมายคล้ายกันอยู่ใกล้กัน จึงค้น “ความหมาย” ได้ดีกว่า\n'
                           'keyword บางกรณี\n'
                           '\n'
                           '18.2 Similarity\n'
                           'Vector search ใช้ similarity metric เช่น cosine similarity เพื่อหา chunks '
                           'ที่ใกล้ query. TPM ไม่ต้องคำนวณเอง แต่\n'
                           'ต้องรู้ว่า top-k และ threshold กระทบ recall/noise\n'
                           '\n'
                           '18.3 Chunking\n'
                           'เอกสารต้องแบ่งเป็น chunks. เล็กเกินสูญ context; ใหญ่เกินดึง noise/ใช้ token '
                           'มาก. ควรใช้ heading/semantic\n'
                           'boundaries และ overlap ตามประเภทเอกสาร\n'
                           '\n'
                           '18.4 Metadata filters\n'
                           'Filter เช่น product, country, effective date, document type ช่วย retrieval '
                           'precision และ policy versioning.\n'
                           'สำคัญใน enterprise knowledge\n'
                           '\n'
                           '18.5 Hybrid search\n'
                           'รวม keyword/BM25 กับ vector search เพื่อจับทั้ง exact terms เช่น product '
                           'code และ semantic meaning\n'
                           '\n'
                           '18.6 Retrieval metrics\n'
                           'Recall@k: คำตอบที่ถูกอยู่ใน top-k หรือไม่; precision/relevance ของ '
                           'retrieved chunks; MRR/nDCG ใช้เมื่อ\n'
                           'ranking สำคัญ\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ระบบใช้ embedding model ตัวไหน | 系统使用哪个嵌入模型？\n'
                           '\n'
                           'เราจะแบ่งเอกสารเป็น chunk อย่างไร | 我们怎么切分文档？\n'
                           '\n'
                           'กรุณาเพิ่ม metadata filter | 请增加元数据过滤条件。\n'
                           '\n'
                           'Retrieval ดึงข้อมูลที่เกี่ยวข้องหรือไม่ | 检索结果是否相关？\n'
                           '\n'
                           'เราจะใช้ hybrid search | 我们会使用混合检索。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Embedding 嵌入\n'
                           '\uf0b7 | Vector 向量\n'
                           '\uf0b7 | Retrieval 检索\n'
                           '\uf0b7 | Metadata 元数据\n'
                           '\uf0b7 | Filter 过滤\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. Chunk เล็ก/ใหญ่เกินไปมีผลอย่างไร\n'
                           '\n'
                           '2. Metadata filter ช่วย policy knowledge อย่างไร\n'
                           '\n'
                           '3. Keyword search ยังมีประโยชน์เมื่อไร\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-19',
                'chapter_title': 'Day 19 — RAG Architecture',
                'order': 19,
                'content': 'Learning Objective\n'
                           ' ออกแบบ Retrieval-Augmented Generation และรู้จุด failure ของแต่ละ stage\n'
                           '\n'
                           '19.1 RAG flow\n'
                           'Ingestion: อ่านเอกสาร -> clean -> chunk -> embedding -> vector index. '
                           'Query: user question -> optional\n'
                           'rewrite -> retrieve -> rerank/filter -> build prompt -> LLM -> answer + '
                           'citation -> feedback/logging\n'
                           '\n'
                           ' INGESTION\n'
                           ' Document -> Clean -> Chunk -> Embed -> Vector Index\n'
                           '\n'
                           ' QUERY\n'
                           ' Question -> Retrieve -> Rerank -> Prompt + Context -> LLM -> Answer + '
                           'Citation\n'
                           '\n'
                           '19.2 Why RAG\n'
                           'ใช้ knowledge ภายใน/อัปเดตบ่อยโดยไม่ retrain model ทั้งตัว, ทำ '
                           'citation/traceability ได้, ควบคุม source ได้มาก\n'
                           'ขึ้น แต่ RAG ไม่การันตีไม่มี hallucination\n'
                           '\n'
                           '19.3 Retrieval failure vs Generation failure\n'
                           'ถ้าดึงเอกสารผิด ต่อให้ LLM เก่งก็อาจตอบผิด. Evaluation ต้องแยก retrieval '
                           'quality กับ answer generation เพื่อ\n'
                           'debug ได้\n'
                           '\n'
                           '19.4 Freshness & versioning\n'
                           'Policy มี effective date ต้อง ensure index อัปเดตและเลิกใช้ version เก่า. '
                           'Ingestion pipeline ต้อง monitor failed\n'
                           'documents และ stale index\n'
                           '\n'
                           '19.5 Access-aware RAG\n'
                           'User ไม่ควร retrieve เอกสารที่ไม่มีสิทธิ์ แม้ final answer ถูก mask ภายหลัง '
                           'ควร enforce permissions ตั้งแต่\n'
                           'retrieval layer\n'
                           '\n'
                           '19.6 Citation\n'
                           'Citation ควรชี้ source/chunk ที่รองรับ claim จริง '
                           'ไม่ใช่แค่แนบเอกสารใกล้เคียง\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ระบบต้องค้นเอกสารที่เกี่ยวข้องก่อน | 系统需要先检索相关文件。\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ข้อมูลนี้มาจาก knowledge base | 这个信息来自知识库。\n'
                           '\n'
                           'ผล retrieval ไม่เกี่ยวข้องกับคำถาม | 检索结果和问题不相关。\n'
                           '\n'
                           'กรุณาแสดง citation ของคำตอบ | 请显示答案的引用来源。\n'
                           '\n'
                           'Index นี้อัปเดตครั้งล่าสุดเมื่อไร | 这个索引最后更新时间是什么时候？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | RAG 检索增强生成\n'
                           '\uf0b7 | Knowledge base 知识库\n'
                           '\uf0b7 | Citation 引用\n'
                           '\uf0b7 | Index 索引\n'
                           '\uf0b7 | Document 文件\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. อธิบาย RAG flow ตั้งแต่ ingestion ถึง answer\n'
                           '\n'
                           '2. RAG ลด hallucination แต่ไม่กำจัด เพราะอะไร\n'
                           '\n'
                           '3. Access-aware retrieval สำคัญอย่างไร\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-20',
                'chapter_title': 'Day 20 — Fine-tuning / RAG / Prompt / Model Choice',
                'order': 20,
                'content': 'Learning Objective\n'
                           ' เลือกวิธีปรับระบบ AI ตามปัญหา ไม่ใช้เทคนิคผิดประเภท\n'
                           '\n'
                           '20.1 Prompt first\n'
                           'ถ้าปัญหาคือ instruction/format ให้ลอง prompt/schema/examples '
                           'ก่อนเพราะเร็วและถูก\n'
                           '\n'
                           '20.2 RAG for knowledge\n'
                           'ถ้าปัญหาคือ model ไม่รู้ข้อมูลบริษัท/ข้อมูลเปลี่ยนบ่อย ใช้ RAG เพราะ update '
                           'knowledge source ได้โดยไม่ retrain\n'
                           '\n'
                           '20.3 Fine-tuning for behavior\n'
                           'เหมาะเมื่ออยากปรับ style/task pattern/structured behavior จาก examples '
                           'จำนวนมาก แต่ไม่ใช่วิธีดีที่สุดสำหรับ\n'
                           'factual knowledge ที่เปลี่ยนรายวัน\n'
                           '\n'
                           '20.4 Model selection\n'
                           'ดู quality บน use case จริง, latency, cost, context, language, tool use, '
                           'privacy, deployment region, SLA,\n'
                           'vendor lock-in, stability. Benchmark ด้วย evaluation set ของตัวเอง\n'
                           '\n'
                           '20.5 Small vs large models\n'
                           'Model ใหญ่ไม่ได้ชนะทุก use case. Routing อาจส่ง simple task ไป model '
                           'เล็กและ difficult task ไป model ใหญ่ ลด\n'
                           'cost/latency\n'
                           '\n'
                           '20.6 Build vs buy\n'
                           'ซื้อ API เร็วแต่ vendor dependency; self-host ควบคุมสูงแต่ต้องมี infra/ops. '
                           'Decision ต้องรวม total cost of\n'
                           'ownership ไม่ใช่ token price อย่างเดียว\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'กรณีนี้ควรใช้ RAG หรือ fine-tuning | 这个场景应该使用 RAG 还是微调？\n'
                           '\n'
                           'เราต้อง benchmark หลาย model | 我们需要对多个模型做基准测试。\n'
                           '\n'
                           'Model ใหญ่มี cost สูงกว่า | 大模型的成本更高。\n'
                           '\n'
                           'ข้อมูลเปลี่ยนบ่อยจึงเหมาะกับ RAG | 数据经常变化，所以更适合 RAG。\n'
                           '\n'
                           'เราต้องประเมิน vendor lock-in | 我们需要评估供应商锁定风险。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Fine-tuning 微调\n'
                           '\uf0b7 | Benchmark 基准测试\n'
                           '\uf0b7 | Cost 成本\n'
                           '\uf0b7 | Vendor 供应商\n'
                           '\uf0b7 | Lock-in 锁定\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. เมื่อไร prompt/RAG/fine-tuning เหมาะคนละแบบ\n'
                           '\n'
                           '2. Model selection ต้องดูอะไรนอกจาก quality\n'
                           '\n'
                           '3. Build vs buy มี trade-off อะไร\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-21',
                'chapter_title': 'Day 21 — Evaluation / Hallucination / Safety',
                'order': 21,
                'content': 'Learning Objective\n'
                           ' สร้าง evaluation system เพื่อรู้ว่า AI ดีพอสำหรับ production หรือยัง\n'
                           '\n'
                           '21.1 Evaluation dataset\n'
                           'สร้างชุดคำถาม representative ครอบคลุม common, edge, adversarial, '
                           'multilingual และ high-risk cases พร้อม\n'
                           'expected answer/source/rubric\n'
                           '\n'
                           '21.2 Dimensions\n'
                           'Correctness, groundedness, relevance, completeness, citation accuracy, '
                           'instruction following, safety,\n'
                           'latency, cost. Metric ต้องสัมพันธ์กับ user/business risk\n'
                           '\n'
                           '21.3 Human eval + automated eval\n'
                           'Human review แม่นด้าน nuance แต่แพง/ช้า; automated/LLM-as-judge scalable '
                           'แต่มี bias. ใช้ผสมและ calibrate\n'
                           'กับ human labels\n'
                           '\n'
                           '21.4 Hallucination\n'
                           'คำตอบดูสมเหตุผลแต่ unsupported/ผิด. ลดด้วย RAG, constrained output, tool '
                           'validation, temperature,\n'
                           'refusal when evidence insufficient และ human review แต่ต้องวัดจริง\n'
                           '\n'
                           '21.5 Guardrails\n'
                           'Input/output filters, policy checks, PII masking, tool permission, rate '
                           'limit, safe fallback. Guardrail เป็นชั้น\n'
                           'เพิ่ม ไม่แทน secure architecture\n'
                           '\n'
                           '21.6 Regression\n'
                           'ทุก model/prompt/retrieval change ต้องรัน benchmark เดิมเพื่อดูว่า quality '
                           'ดีขึ้นส่วนหนึ่งแต่พังอีกส่วนหรือไม่\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'คำตอบนี้ถูกต้องหรือไม่ | 这个回答准确吗？\n'
                           '\n'
                           'เราต้องประเมินคุณภาพของ model | 我们需要评估模型的质量。\n'
                           '\n'
                           'คำตอบนี้มีหลักฐานรองรับหรือไม่ | 这个回答有依据吗？\n'
                           '\n'
                           ' Hallucination rate ต้องต่ำกว่าสอง\n'
                           '                                       幻觉率必须低于百分之二。\n'
                           ' เปอร์เซ็นต์\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           '    กรณีไม่มีข้อมูลเพียงพอ model ต้อง\n'
                           '                                          信息不足时，模型应该拒绝回答。\n'
                           '    ปฏิเสธ\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Evaluation 评估\n'
                           '\uf0b7 | Accuracy 准确率\n'
                           '\uf0b7 | Hallucination 幻觉\n'
                           '\uf0b7 | Evidence 依据\n'
                           '\uf0b7 | Refuse 拒绝\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. สร้าง evaluation dimensions 6 ตัวสำหรับ RAG bot\n'
                           '\n'
                           '2. ทำไม LLM-as-judge ต้อง calibrate\n'
                           '\n'
                           '3. เขียน fallback เมื่อ evidence ไม่พอ\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่\n'
                           '\n'
                           'WEEK 4 - AI TECHNICAL PRODUCT &\n'
                           '                 PLATFORM\n'
                           '     จาก AI Demo สู่ Platform/Production Product'},
               {'chapter_id': 'tpm-day-22',
                'chapter_title': 'Day 22 — AI Product Discovery & Use-case Selection',
                'order': 22,
                'content': 'Learning Objective\n'
                           ' เลือก AI use case ที่มี value และควบคุม risk ได้\n'
                           '\n'
                           '22.1 AI use-case scoring\n'
                           'ให้คะแนน Value, Feasibility, Data readiness, Risk, Adoption friction. Use '
                           'case ที่ value สูงแต่ risk/ข้อมูลไม่\n'
                           'พร้อมอาจเริ่มด้วย assistive mode ก่อน autonomous\n'
                           '\n'
                           '22.2 Assist vs Automate\n'
                           'Assist: AI เสนอคำตอบให้คนตัดสินใจ; Automate: AI execute action. ยิ่ง '
                           'automation สูงต้องเพิ่ม confidence\n'
                           'threshold, approval, audit และ rollback\n'
                           '\n'
                           '22.3 Human-in-the-loop\n'
                           'กำหนดว่าเคสไหนต้อง human review เช่น low confidence, high amount, policy '
                           'exception, sensitive content.\n'
                           'HITL ต้องออกแบบ queue/SLA ไม่ใช่ใส่ชื่อไว้เฉย ๆ\n'
                           '\n'
                           '22.4 Adoption\n'
                           'Internal AI tool สำเร็จไม่ใช่แค่ model score ต้อง integrate workflow, ลด '
                           'clicks, training, trust/citation,\n'
                           'feedback loop และ champion network\n'
                           '\n'
                           '22.5 North Star + guardrails\n'
                           'North star เช่น time-to-answer ลดลง; guardrail เช่น incorrect-answer rate '
                           'ไม่เกิน 1%, privacy incident = 0\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Use case นี้สร้าง value อะไร | 这个应用场景能带来什么价值？\n'
                           '\n'
                           'เคสนี้ต้องมี human review | 这个场景需要人工审核。\n'
                           '\n'
                           'เราควรเริ่มจาก assistive mode ก่อน | 我们应该先从辅助模式开始。\n'
                           '\n'
                           'Adoption ของผู้ใช้เป็นอย่างไร | 用户采用率怎么样？\n'
                           '\n'
                           'Guardrail metric ของเราคืออะไร | 我们的护栏指标是什么？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Value 价值\n'
                           '\uf0b7 | Human review 人工审核\n'
                           '\uf0b7 | Adoption 采用率\n'
                           '\uf0b7 | Guardrail 护栏\n'
                           '\uf0b7 | Assist 辅助\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. ให้คะแนน AI use case หนึ่งตัว 5 มิติ\n'
                           '\n'
                           '2. Assist กับ automate ต่างกันด้าน risk อย่างไร\n'
                           '\n'
                           '3. ตั้ง north star + guardrails สำหรับ internal copilot\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-23',
                'chapter_title': 'Day 23 — AI Architecture & End-to-End Design',
                'order': 23,
                'content': 'Learning Objective\n'
                           ' ออกแบบ AI application เป็นระบบ production ไม่ใช่ demo\n'
                           '\n'
                           '23.1 Reference architecture\n'
                           'User -> UI -> API/Auth -> Orchestrator -> Retrieval/Tools -> Model Gateway '
                           '-> LLM ->\n'
                           'Post-processing/Guardrails -> Response. ด้านข้างมี Logging, Evaluation, '
                           'Cost tracking, Feedback, Secrets,\n'
                           'Policy\n'
                           '\n'
                           ' User/UI\n'
                           '   |\n'
                           ' API + Auth\n'
                           '   |\n'
                           ' AI Orchestrator\n'
                           '   |------ Retrieval / Vector DB\n'
                           '   |------ Tools / Enterprise APIs\n'
                           '   |\n'
                           ' Model Gateway -> LLM\n'
                           '   |\n'
                           ' Guardrails + Citation -> Response\n'
                           '\n'
                           ' Cross-cutting: Logging | Evaluation | Cost | Security\n'
                           '\n'
                           '23.2 Orchestration\n'
                           'Orchestrator ตัดสินใจว่าจะ retrieve, call tool, route model, retry, '
                           'validate output อย่างไร. Logic สำคัญควร\n'
                           'deterministic/observable ไม่ฝากทุกอย่างให้ model\n'
                           '\n'
                           '23.3 Tool calling\n'
                           'LLM อาจเรียก API เช่นค้นยอด/สร้าง ticket. ต้องกำหนด schema, allowlist, '
                           'authorization, confirmation สำหรับ\n'
                           'destructive action และ idempotency\n'
                           '\n'
                           '23.4 Model gateway\n'
                           'ชั้นกลางช่วย route provider/model, enforce policies, logging, quota, '
                           'fallback และ reduce lock-in. มีประโยชน์\n'
                           'เมื่อองค์กรมีหลายทีม/หลาย model\n'
                           '\n'
                           '23.5 Failure modes\n'
                           'Model timeout, rate limit, retrieval empty, malformed JSON, tool error, '
                           'provider outage. ทุก failure ต้องมี\n'
                           'timeout/retry/fallback/user message ที่ตั้งใจออกแบบ\n'
                           '\n'
                           '23.6 Data boundary\n'
                           'วาดว่าข้อมูล sensitive ออกจาก trust boundary ตรงไหน และ '
                           'encryption/retention/provider policy เป็นอย่างไร\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ช่วยอธิบาย end-to-end architecture | 请解释一下端到端架构。\n'
                           '\n'
                           'Model เรียก tool อะไรได้บ้าง | 模型可以调用哪些工具？\n'
                           '\n'
                           '                                             如果供应商服务不可用，有备用模型\n'
                           '    ถ้า provider ล่มมี fallback model ไหม\n'
                           '                                             吗？\n'
                           '\n'
                           'Output ต้องผ่าน validation ก่อน | 输出必须先通过验证。\n'
                           '\n'
                           '    ข้อมูล sensitive ถูกส่งออกนอกระบบ\n'
                           '                                             敏感数据会发送到系统外部吗？\n'
                           '    หรือไม่\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | End-to-end 端到端\n'
                           '\uf0b7 | Tool 工具\n'
                           '\uf0b7 | Provider 供应商\n'
                           '\uf0b7 | Validation 验证\n'
                           '\uf0b7 | Sensitive 敏感\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. วาด reference architecture 10 components\n'
                           '\n'
                           '2. Tool calling ต้องมี control อะไร\n'
                           '\n'
                           '3. ยก failure mode 5 ตัวพร้อม fallback\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-24',
                'chapter_title': 'Day 24 — Model Economics / Latency / Cost / Scale',
                'order': 24,
                'content': 'Learning Objective\n'
                           ' คำนวณ unit economics และออกแบบ scale decision ระดับ Product\n'
                           '\n'
                           '24.1 Cost drivers\n'
                           'AI cost อาจมาจาก input/output tokens, embedding, vector storage/search, '
                           'reranking, tool APIs, compute,\n'
                           'observability และ human review. อย่าดู token price อย่างเดียว\n'
                           '\n'
                           '24.2 Unit economics\n'
                           'Cost per request x requests/day x days/month = base variable cost. เพิ่ม '
                           'peak capacity, cache hit, retry rate,\n'
                           'long-context users และ support overhead เพื่อ scenario planning\n'
                           '\n'
                           '24.3 Latency budget\n'
                           'แบ่ง latency end-to-end: auth 100ms + retrieval 400ms + rerank 300ms + LLM '
                           '2.5s + postprocess 200ms =\n'
                           '~3.5s. ถ้าช้าให้ optimize component ที่กิน budget มากที่สุด\n'
                           '\n'
                           '24.4 P50/P95/P99\n'
                           'Average ซ่อน tail latency. P95 = 95% requests เร็วกว่าค่านี้. User '
                           'experience มักเสียจาก tail จึงควรตั้ง SLO เป็น\n'
                           'percentile\n'
                           '\n'
                           '24.5 Caching\n'
                           'Cache embedding/retrieval/response สำหรับ query ซ้ำบางกรณี ลด cost/latency '
                           'แต่ต้องคิด\n'
                           'privacy/freshness/invalidation\n'
                           '\n'
                           '24.6 Scale 1000+ users\n'
                           'ต้องคิด concurrency, quota, rate limit, tenancy, permissions, rollout, '
                           'support, model capacity และ cost\n'
                           'allocation per team\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Cost ต่อ request เท่าไร | 每个请求的成本是多少？\n'
                           '\n'
                           'P95 latency เป้าหมายคือเท่าไร | P95 延迟目标是多少？\n'
                           '\n'
                           ' ช่วง peak รองรับ concurrent users\n'
                           '                                       高峰期可以支持多少并发用户？\n'
                           ' เท่าไร\n'
                           '\n'
                           'เราต้องลด token usage | 我们需要降低 token 使用量。\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'Cache จะกระทบ freshness หรือไม่ | 缓存会影响数据新鲜度吗？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Cost 成本\n'
                           '\uf0b7 | Latency 延迟\n'
                           '\uf0b7 | Concurrent 并发\n'
                           '\uf0b7 | Cache 缓存\n'
                           '\uf0b7 | Peak 高峰\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. คำนวณ cost/month ถ้า 20,000 requests/day x 0.40 บาท\n'
                           '\n'
                           '2. P95 ต่างจาก average อย่างไร\n'
                           '\n'
                           '3. วาด latency budget ของ RAG flow\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-25',
                'chapter_title': 'Day 25 — MLOps / LLMOps / Model Lifecycle',
                'order': 25,
                'content': 'Learning Objective\n'
                           ' เข้าใจ lifecycle, versioning, deployment, monitoring, evaluation ตาม JD '
                           'สาย AI platform\n'
                           '\n'
                           '25.1 Lifecycle\n'
                           'Experiment -> Evaluate -> Approve -> Package -> Deploy -> Observe -> '
                           'Compare -> Rollback/Promote ->\n'
                           'Retire. Artifacts ที่ version ต้องมี model, prompt, dataset, retrieval '
                           'config, code, evaluation results\n'
                           '\n'
                           '25.2 Model registry\n'
                           'เก็บ model versions, metadata, lineage, metrics, approval stage ช่วย '
                           'governance/reproducibility. สำหรับ\n'
                           'hosted models อาจ registry configuration/provider version แทน binary model\n'
                           '\n'
                           '25.3 Experiment tracking\n'
                           'บันทึกว่า run ไหนใช้ dataset/prompt/model/config อะไรและได้ metric เท่าไร '
                           'เพื่อเปรียบเทียบแบบ reproducible\n'
                           '\n'
                           '25.4 Online monitoring\n'
                           'Monitor system metrics + AI metrics: latency/errors/token/cost, drift, '
                           'retrieval quality proxy, refusal, safety\n'
                           'flags, user feedback. Ground truth บางอย่างมาทีหลัง จึงต้องมี delayed '
                           'evaluation\n'
                           '\n'
                           '25.5 Canary / Shadow / A-B\n'
                           'Canary ปล่อยให้ user บางส่วน; shadow ส่ง traffic copy ให้ model '
                           'ใหม่แต่ไม่แสดงผล; A/B เปรียบ\n'
                           'behavior/outcome. เลือกตาม risk\n'
                           '\n'
                           '25.6 Rollback & version compatibility\n'
                           'Model/prompt/retriever versions ต้อง trace response ได้. ถ้าคุณไม่ตอบได้ว่า '
                           '“คำตอบนี้มาจาก version ไหน”\n'
                           'governance ยังไม่พร้อม\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ตอนนี้ใช้ model version ไหน | 现在使用哪个模型版本？\n'
                           '\n'
                           'เราต้องบันทึก experiment ทุกครั้ง | 我们需要记录每次实验。\n'
                           '\n'
                           ' Model ใหม่ต้องผ่าน regression test\n'
                           '                                       新模型需要先通过回归测试。\n'
                           ' ก่อน\n'
                           '\n'
                           'เราจะปล่อยแบบ shadow ก่อน | 我们先进行影子测试。\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'คำตอบนี้มาจาก model version ไหน | 这个回答来自哪个模型版本？\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Model version 模型版本\n'
                           '\uf0b7 | Experiment 实验\n'
                           '\uf0b7 | Regression 回归\n'
                           '\uf0b7 | Shadow test 影子测试\n'
                           '\uf0b7 | Lifecycle 生命周期\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. อธิบาย model lifecycle 8 stages\n'
                           '\n'
                           '2. Canary/shadow/A-B ต่างกันอย่างไร\n'
                           '\n'
                           '3. ทำไม lineage สำคัญต่อ regulated industry\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-26',
                'chapter_title': 'Day 26 — Platform Product / Internal Developer Tools',
                'order': 26,
                'content': 'Learning Objective\n'
                           '  เข้าใจ Platform Product ซึ่งมักเป็นแกนของ Technical PM ระดับสูง\n'
                           '\n'
                           '26.1 Platform as product\n'
                           'Platform ให้ capabilities reusable แก่ internal developers/teams เช่น API '
                           'platform, model gateway, feature\n'
                           'store, deployment platform. “User” คือ developer/data scientist/product '
                           'team\n'
                           '\n'
                           '26.2 Self-service\n'
                           'เป้าหมายคือให้ทีมทำงานเองผ่าน portal/API/SDK โดยมี guardrails ไม่ต้องเปิด '
                           'ticket ทุกเรื่อง. Metric เช่น time-to-\n'
                           'first-success, setup lead time, deployment frequency\n'
                           '\n'
                           '26.3 Golden path\n'
                           'สร้างวิธีมาตรฐานที่ง่ายและปลอดภัย เช่น template service + logging/security '
                           'CI pipeline พร้อมใช้ แต่อนุญาต\n'
                           'escape hatch สำหรับ case พิเศษ\n'
                           '\n'
                           '26.4 Developer Experience\n'
                           'Docs, SDK, examples, error messages, sandbox, observability และ support '
                           'เป็น product experience.\n'
                           'Platform ที่ technically powerful แต่ใช้ยาก adoption ต่ำ\n'
                           '\n'
                           '26.5 Multi-tenancy\n'
                           'หลายทีมใช้ platform เดียว ต้องแยก quota, cost, permissions, data, '
                           'environments และ noisy-neighbor risk\n'
                           '\n'
                           '26.6 Platform roadmap\n'
                           'Prioritize common pain across teams, not one-off request จากทีมเสียงดัง. '
                           'ใช้ adoption, reuse, time saved,\n'
                           'reliability และ satisfaction เป็น metrics\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ผู้ใช้ของ platform คือ developer | 平台的用户是开发人员。\n'
                           '\n'
                           'เราต้องทำ self-service ให้มากขึ้น | 我们需要提高自助服务能力。\n'
                           '\n'
                           ' Developer ใช้เวลานานแค่ไหนกว่าจะ\n'
                           '                                       开发人员需要多长时间才能开始使用？\n'
                           ' เริ่มได้\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'เราต้องมี standard template | 我们需要提供标准模板。\n'
                           '\n'
                           'แต่ละทีมมี quota แยกกัน | 每个团队都有独立配额。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Platform 平台\n'
                           '\uf0b7 | Developer 开发人员\n'
                           '\uf0b7 | Self-service 自助服务\n'
                           '\uf0b7 | Template 模板\n'
                           '\uf0b7 | Quota 配额\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. Platform product user แตกต่างจาก consumer app อย่างไร\n'
                           '\n'
                           '2. Golden path คืออะไร\n'
                           '\n'
                           '3. ตั้ง platform metrics 5 ตัว\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-27',
                'chapter_title': 'Day 27 — Metrics / Experimentation / Rollout',
                'order': 27,
                'content': 'Learning Objective\n'
                           ' เชื่อม product outcome กับ technical/AI metrics และ rollout '
                           'ที่เรียนรู้ได้\n'
                           '\n'
                           '27.1 Metric stack\n'
                           'Business: revenue/cost/risk; Product: adoption/task success/time saved; AI: '
                           'quality/groundedness; System:\n'
                           'latency/availability/error; Cost: unit economics. ต้องเห็น causal chain\n'
                           '\n'
                           '27.2 Leading vs Lagging\n'
                           'Leading เช่น weekly active users/feedback; lagging เช่น productivity/cost '
                           'saving. อย่ารอ lagging metric อย่าง\n'
                           'เดียว\n'
                           '\n'
                           '27.3 Experiment design\n'
                           'Define hypothesis, population, control/treatment, primary metric, '
                           'guardrails, duration, sample. AI\n'
                           'experiments ต้อง freeze/version configs เพื่อ interpret results\n'
                           '\n'
                           '27.4 Rollout stages\n'
                           'Internal dogfood -> pilot -> limited GA -> wider rollout -> full. แต่ละ '
                           'stage มี exit criteria เช่น quality,\n'
                           'incident rate, support volume\n'
                           '\n'
                           '27.5 Feedback loop\n'
                           'Thumbs up/down อย่างเดียวไม่พอ; capture reason categories, correction, '
                           'source issue และ task outcome เพื่อ\n'
                           'ปรับ retrieval/prompt/product\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'North Star metric คืออะไร | 北极星指标是什么？\n'
                           '\n'
                           'เราต้องตั้ง guardrail metrics | 我们需要设置护栏指标。\n'
                           '\n'
                           'Pilot จะเริ่มกับผู้ใช้หนึ่งร้อยคน | 试点将从一百名用户开始。\n'
                           '\n'
                           'ผลการทดลองมีนัยสำคัญหรือไม่ | 实验结果是否显著？\n'
                           '\n'
                           ' เราต้องเก็บเหตุผลของ negative\n'
                           '                                        我们需要记录负面反馈的原因。\n'
                           ' feedback\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Metric 指标\n'
                           '\uf0b7 | Pilot 试点\n'
                           '\uf0b7 | Experiment 实验\n'
                           '\uf0b7 | Feedback 反馈\n'
                           '\uf0b7 | Significant 显著\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. สร้าง metric stack สำหรับ AI assistant\n'
                           '\n'
                           '2. Rollout stage ควรมี exit criteria อะไร\n'
                           '\n'
                           '3. Feedback แบบไหนช่วย debug RAG ได้จริง\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-28',
                'chapter_title': 'Day 28 — AI Production Incident / Governance',
                'order': 28,
                'content': 'Learning Objective\n'
                           ' รับมือ AI incident และ governance ในองค์กรจริง\n'
                           '\n'
                           '28.1 AI-specific incident\n'
                           'ตัวอย่าง: hallucination เพิ่มหลังเปลี่ยน model, prompt injection ทำให้ leak '
                           'data, retrieval index stale, provider\n'
                           'outage, cost spike, toxic output. Severity ต้องรวม customer/regulatory '
                           'impact\n'
                           '\n'
                           '28.2 Containment\n'
                           'Feature flag off, route fallback model, disable tool action, restrict '
                           'high-risk users, freeze ingestion, revert\n'
                           'prompt/model. ต้องมี runbook ล่วงหน้า\n'
                           '\n'
                           '28.3 Investigation\n'
                           'Trace request -> model/prompt/retrieval/tool version -> logs -> source docs '
                           '-> provider status. ถ้า lineage ไม่\n'
                           'ครบ debug ช้า\n'
                           '\n'
                           '28.4 Governance\n'
                           'Model inventory, use-case risk tier, approval, evaluation evidence, data '
                           'classification, access, change\n'
                           'management, audit, monitoring และ retirement เป็น lifecycle governance\n'
                           '\n'
                           '28.5 Responsible AI\n'
                           'Fairness, transparency, privacy, explainability ตาม use case. High-stakes '
                           'decision ควรมี clear\n'
                           'accountability และ human oversight\n'
                           '\n'
                           '28.6 Postmortem\n'
                           'Blameless: timeline, impact, root cause, contributing factors, detection '
                           'gap, corrective actions,\n'
                           'owners/deadlines, what went well/poorly\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'AI incident นี้มีผลกระทบระดับไหน | 这次 AI 事故的影响等级是什么？\n'
                           '\n'
                           'เราต้องปิด feature ชั่วคราว | 我们需要暂时关闭这个功能。\n'
                           '\n'
                           ' กรุณาตรวจสอบ model/prompt\n'
                           '                                      请检查模型和提示词版本。\n'
                           ' version\n'
                           '\n'
                           'ปัญหานี้เกี่ยวกับข้อมูลหรือ model | 这个问题和数据有关还是和模型有关？\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           '    เราต้องมี corrective action ที่มี\n'
                           '                                         我们需要有明确负责人的整改措施。\n'
                           '    owner\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Incident 事故\n'
                           '\uf0b7 | Impact 影响\n'
                           '\uf0b7 | Corrective action 整改措施\n'
                           '\uf0b7 | Responsible person 负责人\n'
                           '\uf0b7 | Review 复盘\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. เขียน containment 5 ตัวสำหรับ AI incident\n'
                           '\n'
                           '2. Governance artifacts ที่ต้องมีอะไรบ้าง\n'
                           '\n'
                           '3. ทำไม lineage สำคัญในการ investigate\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-29',
                'chapter_title': 'Day 29 — Chinese Vendor Technical Meeting',
                'order': 29,
                'content': 'Learning Objective\n'
                           ' ใช้ภาษาจีนประชุม requirement, defect, timeline, architecture และ '
                           'production issue\n'
                           '\n'
                           '29.1 Meeting structure\n'
                           'เปิด meeting ด้วย objective + decisions needed. ระหว่างคุยถาม fact ก่อน '
                           'opinion: current behavior, expected\n'
                           'behavior, logs, version, environment, reproduction steps, impact. ปิดด้วย '
                           'action/owner/date\n'
                           '\n'
                           '29.2 Clarification language\n'
                           'ใช้ประโยคสั้น ชัด ไม่แปลศัพท์ technical ทุกคำเป็นจีนก็ได้ เพราะ vendor '
                           'จีนใช้ API, JSON, log, bug, release, server\n'
                           'ปนอังกฤษบ่อย เป้าหมายคือสื่อสารงาน ไม่ใช่สอบภาษา\n'
                           '\n'
                           '29.3 Defect discussion\n'
                           'ต้องแยก expected vs actual, reproducible steps, environment/version, data, '
                           'severity, workaround, root\n'
                           'cause, ETA, regression scope\n'
                           '\n'
                           '29.4 Architecture discussion\n'
                           'ถาม data flow, interface, auth, dependency, sync/async, timeout/retry, '
                           'fallback, performance, security,\n'
                           'monitoring\n'
                           '\n'
                           '29.5 Negotiation\n'
                           'เมื่อ ETA ไม่ทัน ให้ถาม option: ลด scope, parallelize, workaround, phase '
                           'delivery, additional resource และ\n'
                           'impact ของแต่ละทาง ไม่ใช่ถาม “ทำให้ทันได้ไหม” อย่างเดียว\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'วันนี้เราต้องตัดสินใจสามเรื่อง | 今天我们需要决定三件事。\n'
                           '\n'
                           ' ช่วยอธิบาย current behavior กับ\n'
                           '                                       请说明当前行为和预期行为。\n'
                           ' expected behavior\n'
                           '\n'
                           'ปัญหานี้ reproduce ได้ไหม | 这个问题可以重现吗？\n'
                           '\n'
                           'กรุณาส่ง log และ request ID | 请发送日志和请求 ID。\n'
                           '\n'
                           'Root cause ยืนยันแล้วหรือยัง | 根本原因已经确认了吗？\n'
                           '\n'
                           'ทางเลือกอื่นมีอะไรบ้าง | 还有哪些替代方案？\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ถ้าลด scope จะส่งได้เร็วขึ้นเท่าไร | 如果缩小范围，可以提前多久交付？\n'
                           '\n'
                           'ใครเป็น owner ของ action นี้ | 谁负责这个行动项？\n'
                           '\n'
                           'กรุณายืนยันวันส่งมอบ | 请确认交付日期。\n'
                           '\n'
                           'เราจะ follow up พรุ่งนี้ | 我们明天再跟进。\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Decision 决定\n'
                           '\uf0b7 | Expected 预期\n'
                           '\uf0b7 | Reproduce 重现\n'
                           '\uf0b7 | Alternative 替代方案\n'
                           '\uf0b7 | Follow up 跟进\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. จำลอง defect meeting 5 นาที\n'
                           '\n'
                           '2. ตั้งคำถาม architecture review ภาษาจีน 5 ข้อ\n'
                           '\n'
                           '3. ปิด meeting ด้วย action/owner/date เป็นภาษาจีน\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-day-30',
                'chapter_title': 'Day 30 — Capstone + Interview Readiness',
                'order': 30,
                'content': 'Learning Objective\n'
                           ' รวมทุกอย่างเป็น project story และเตรียมตอบ interview Technical Product '
                           'Manager\n'
                           '\n'
                           '30.1 Capstone\n'
                           'ออกแบบ AI Banking Knowledge Assistant: Problem = staff ค้น policy ช้า. '
                           'Users = loan officers. Goal = time-\n'
                           'to-answer <2 นาที, citation accuracy >=98%, incorrect answer <1%, P95 '
                           'latency <5s\n'
                           '\n'
                           '30.2 Architecture\n'
                           'Web UI -> SSO/Auth -> API -> AI Orchestrator -> Access-aware Retrieval -> '
                           'Vector DB -> Reranker -> Model\n'
                           'Gateway -> LLM -> Citation/Guardrails -> Response. '
                           'Logging/evaluation/cost/feedback ครอบทุก request\n'
                           '\n'
                           '30.3 Product decisions\n'
                           'RAG แทน fine-tuning สำหรับ policy freshness, metadata filter ตาม '
                           'product/effective date, human review\n'
                           'สำหรับ high-risk question, canary rollout 100 users, fallback to keyword '
                           'search/manual source\n'
                           '\n'
                           '30.4 Production plan\n'
                           'Evaluation set 500+ cases, security review, load test, runbook, '
                           'model/prompt versioning, monitoring, cost\n'
                           'alert, incident owner, hypercare. เปลี่ยน model ต้อง regression ก่อน '
                           'promote\n'
                           '\n'
                           '30.5 Interview framing\n'
                           'ตอบแบบ Situation -> Problem -> Decision -> Trade-off -> Technical depth -> '
                           'Result -> Learning. อย่าเพียง list\n'
                           'tools; อธิบายว่าทำไมตัดสินใจและ metric เปลี่ยนอย่างไร\n'
                           '\n'
                           '30.6 Gap honesty\n'
                           'ถ้ายังไม่เคย build ML platform จริง ให้พูดตรง ๆ ว่ามี production product '
                           'experience อะไร แล้วแสดง\n'
                           'portfolio/technical understanding ที่เติมมา หลีกเลี่ยงการ claim hands-on '
                           'engineering ที่ไม่ได้ทำ\n'
                           '\n'
                           'ภาษาจีนสำหรับงานวันนี้\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'ฉันจะอธิบาย architecture นี้แบบ end-\n'
                           '                                        我会从端到端解释这个架构。\n'
                           'to-end\n'
                           '\n'
                           'เราเลือก RAG เพราะข้อมูลเปลี่ยนบ่อย | 我们选择 RAG，因为数据经常变化。\n'
                           '\n'
                           'Metric หลักของเราคือ time-to-\n'
                           '                                        我们的核心指标是回答时间。\n'
                           'answer\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           '    ก่อน go-live ต้องผ่าน evaluation และ\n'
                           '                                            上线前必须通过评估和安全审查。\n'
                           '    security review\n'
                           '\n'
                           '    ทุกการเปลี่ยน model ต้องมี\n'
                           '                                            每次模型变更都需要回归测试。\n'
                           '    regression test\n'
                           '\n'
                           'Key Vocabulary\n'
                           '\uf0b7 | Architecture 架构\n'
                           '\uf0b7 | Core metric 核心指标\n'
                           '\uf0b7 | Security review 安全审查\n'
                           '\uf0b7 | Evaluation 评估\n'
                           '\uf0b7 | Change 变更\n'
                           '\n'
                           'แบบฝึกหัดท้ายบท\n'
                           '1. พูด capstone architecture โดยไม่ดูหนังสือ\n'
                           '\n'
                           '2. ตอบ Why RAG not fine-tuning ภายใน 45 วินาที\n'
                           '\n'
                           '3. ตอบ “How do you manage AI model changes in production?” ภายใน 90 วินาที\n'
                           '\n'
                           '    Explain Back\n'
                           '    ปิดหนังสือ แล้วอธิบายบทวันนี้ด้วยภาษาของตัวเองให้ได้ภายใน 2 นาที '
                           'หากติดศัพท์ ให้กลับไปดูเฉพาะจุดนั้น แล้ว\n'
                           '    อธิบายใหม่'},
               {'chapter_id': 'tpm-appendix-a',
                'chapter_title': 'Appendix A — Technical Chinese Glossary',
                'order': 31,
                'content': 'คำศัพท์ด้านล่างตั้งใจให้ใช้คุยงานจริง '
                           'สามารถใช้คำอังกฤษปนกับภาษาจีนได้ตามธรรมชาติของทีมเทค\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'สถาปัตยกรรม | 架构\n'
                           '\n'
                           'ระบบ | 系统\n'
                           '\n'
                           'บริการ/Service | 服务\n'
                           '\n'
                           'อินเทอร์เฟซ/API | 接口\n'
                           '\n'
                           'ฐานข้อมูล | 数据库\n'
                           '\n'
                           'ข้อมูล | 数据\n'
                           '\n'
                           'ข้อมูลต้นทาง | 数据源\n'
                           '\n'
                           'ฟิลด์ | 字段\n'
                           '\n'
                           'ตารางข้อมูล | 数据表\n'
                           '\n'
                           'แคช | 缓存\n'
                           '\n'
                           'คิว | 队列\n'
                           '\n'
                           'เครือข่าย | 网络\n'
                           '\n'
                           'สิทธิ์ | 权限\n'
                           '\n'
                           'เข้ารหัส | 加密\n'
                           '\n'
                           'ล็อก | 日志\n'
                           '\n'
                           'เฝ้าระวัง | 监控\n'
                           '\n'
                           'แจ้งเตือน | 告警\n'
                           '\n'
                           'ดีพลอย | 部署\n'
                           '\n'
                           'ขึ้นระบบ | 上线\n'
                           '\n'
                           'ย้อนเวอร์ชัน | 回滚\n'
                           '\n'
                           'เวอร์ชัน | 版本\n'
                           '\n'
                           'ทดสอบ | 测试\n'
                           '\n'
                           'ข้อบกพร่อง | 缺陷\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'รีเกรสชัน | 回归\n'
                           '\n'
                           'ประสิทธิภาพ | 性能\n'
                           '\n'
                           'ความหน่วง | 延迟\n'
                           '\n'
                           'ความพร้อมใช้งาน | 可用性\n'
                           '\n'
                           'สเกลได้ | 可扩展性\n'
                           '\n'
                           'ข้อกำหนด | 需求\n'
                           '\n'
                           'ขอบเขต | 范围\n'
                           '\n'
                           'การพัฒนา | 开发\n'
                           '\n'
                           'การส่งมอบ | 交付\n'
                           '\n'
                           'ความเสี่ยง | 风险\n'
                           '\n'
                           'การพึ่งพา | 依赖\n'
                           '\n'
                           'AI | 人工智能\n'
                           '\n'
                           'โมเดล | 模型\n'
                           '\n'
                           'ฝึกโมเดล | 训练\n'
                           '\n'
                           'อินเฟอเรนซ์ | 推理\n'
                           '\n'
                           'เอ็มเบดดิง | 嵌入\n'
                           '\n'
                           'เวกเตอร์ | 向量\n'
                           '\n'
                           'การค้นคืน | 检索\n'
                           '\n'
                           'ฐานความรู้ | 知识库\n'
                           '\n'
                           'ไฟน์จูน | 微调\n'
                           '\n'
                           'ประเมินผล | 评估\n'
                           '\n'
                           'ความแม่นยำ | 准确率\n'
                           '\n'
                           'หลอนข้อมูล | 幻觉\n'
                           '\n'
                           'บริบท | 上下文\n'
                           '\n'
                           'พรอมป์ต์ | 提示词\n'
                           '\n'
                           'ภาษาไทย | 中文\n'
                           '\n'
                           'แพลตฟอร์ม | 平台\n'
                           '\n'
                           'ผู้พัฒนา | 开发人员\n'
                           '\n'
                           'ผู้ให้บริการ/เวนเดอร์ | 供应商\n'
                           '\n'
                           'โควตา | 配额\n'
                           '\n'
                           'เหตุขัดข้อง | 事故\n'
                           '\n'
                           'สาเหตุราก | 根本原因\n'
                           '\n'
                           'ผลกระทบ | 影响\n'
                           '\n'
                           'ตัวชี้วัด | 指标\n'
                           '\n'
                           'ฟีดแบ็ก | 反馈'},
               {'chapter_id': 'tpm-appendix-b',
                'chapter_title': 'Appendix B — 60 คำถามที่ Technical PM ควรถามใน Meeting',
                'order': 32,
                'content': '1. Problem ที่แท้จริงคืออะไร และใครได้รับผลกระทบ?\n'
                           '\n'
                           '2. Outcome ที่ต้องการวัดคืออะไร?\n'
                           '\n'
                           '3. Source of truth ของข้อมูลนี้คือระบบไหน?\n'
                           '\n'
                           '4. User flow เริ่มและจบตรงไหน?\n'
                           '\n'
                           '5. Architecture ปัจจุบันเป็นอย่างไร?\n'
                           '\n'
                           '6. Component ไหนเป็น critical dependency?\n'
                           '\n'
                           '7. API contract เปลี่ยนอะไรบ้าง?\n'
                           '\n'
                           '8. Backward compatibility ยังอยู่หรือไม่?\n'
                           '\n'
                           '9. Authentication และ authorization ใช้วิธีใด?\n'
                           '\n'
                           '10. Timeout ของ API เท่าไร?\n'
                           '\n'
                           '11. Retry policy คืออะไร?\n'
                           '\n'
                           '12. Request นี้ idempotent หรือไม่?\n'
                           '\n'
                           '13. หาก downstream ล่ม fallback คืออะไร?\n'
                           '\n'
                           '14. มี cache หรือไม่ และ stale ได้กี่นาที?\n'
                           '\n'
                           '15. Data consistency เป็น strong หรือ eventual?\n'
                           '\n'
                           '16. มี queue/event ตรงไหน?\n'
                           '\n'
                           '17. Duplicate message จัดการอย่างไร?\n'
                           '\n'
                           '18. Environment ไหน reproduce ปัญหาได้?\n'
                           '\n'
                           '19. Version ที่มีปัญหาคืออะไร?\n'
                           '\n'
                           '20. Request ID / trace ID มีไหม?\n'
                           '\n'
                           '21. Root cause ยืนยันแล้วหรือยัง?\n'
                           '\n'
                           '22. Workaround ตอนนี้คืออะไร?\n'
                           '\n'
                           '23. Permanent fix ต่างจาก workaround อย่างไร?\n'
                           '\n'
                           '24. Impact กี่ users/transactions?\n'
                           '\n'
                           '25. SLO/SLA ที่กระทบคืออะไร?\n'
                           '\n'
                           '26. Monitoring ตอนนี้เห็น signal อะไร?\n'
                           '\n'
                           '27. Alert ทำงานก่อน user report หรือไม่?\n'
                           '\n'
                           '28. Test coverage ของ change นี้มีอะไร?\n'
                           '\n'
                           '29. Regression scope ครอบคลุมอะไร?\n'
                           '\n'
                           '30. Performance target เป็น P95/P99 เท่าไร?\n'
                           '\n'
                           '31. Peak traffic เท่าไร?\n'
                           '\n'
                           '32. Capacity margin เหลือเท่าไร?\n'
                           '\n'
                           '33. Go-live criteria คืออะไร?\n'
                           '\n'
                           '34. Rollback criteria คืออะไร?\n'
                           '\n'
                           '35. ถ้า rollback ไม่ได้ forward-fix plan คืออะไร?\n'
                           '\n'
                           '36. Migration reconciliation ตรวจอะไรบ้าง?\n'
                           '\n'
                           '37. Data loss/duplicate risk มีไหม?\n'
                           '\n'
                           '38. Security review ต้องทำหรือไม่?\n'
                           '\n'
                           '39. PII ออกนอก trust boundary หรือไม่?\n'
                           '\n'
                           '40. Audit log ครบหรือไม่?\n'
                           '\n'
                           '41. AI use case นี้จำเป็นต้องใช้ LLM หรือ rule ก็พอ?\n'
                           '\n'
                           '42. Model นี้ถูกเลือกจาก benchmark อะไร?\n'
                           '\n'
                           '43. Evaluation set สะท้อน production หรือไม่?\n'
                           '\n'
                           '44. Groundedness/citation accuracy เท่าไร?\n'
                           '\n'
                           '45. Hallucination rate เท่าไร?\n'
                           '\n'
                           '46. RAG retrieval fail แยกวัดจาก generation fail หรือไม่?\n'
                           '\n'
                           '47. Index freshness เท่าไร?\n'
                           '\n'
                           '48. Document permissions enforce ตอน retrieval หรือไม่?\n'
                           '\n'
                           '49. Prompt/model/retriever version trace ได้ไหม?\n'
                           '\n'
                           '50. Model provider เก็บ input หรือใช้ train ต่อหรือไม่?\n'
                           '\n'
                           '51. Token cost/request เท่าไร?\n'
                           '\n'
                           '52. P95 end-to-end latency เท่าไร?\n'
                           '\n'
                           '53. Fallback model/experience คืออะไร?\n'
                           '\n'
                           '54. เมื่อ provider rate limit จะทำอย่างไร?\n'
                           '\n'
                           '55. Human review อยู่ตรงไหน?\n'
                           '\n'
                           '56. High-risk cases มี confidence/approval threshold หรือไม่?\n'
                           '\n'
                           '57. AI change ใช้ canary/shadow/A-B แบบไหน?\n'
                           '\n'
                           '58. Metric หลัง rollout คืออะไร?\n'
                           '\n'
                           '59. ใครเป็น owner ของ action นี้?\n'
                           '\n'
                           '60. Decision และ due date คืออะไร?'},
               {'chapter_id': 'tpm-appendix-c',
                'chapter_title': 'Appendix C — Templates ใช้ทำงานจริง',
                'order': 33,
                'content': 'C1. One-page PRD\n'
                           '    1) Problem\n'
                           '    2) User / Job\n'
                           '    3) Objective & Success Metrics\n'
                           '    4) In Scope / Out of Scope\n'
                           '    5) Functional Requirements\n'
                           '    6) Non-functional Requirements\n'
                           '    7) Data & Privacy\n'
                           '    8) Dependencies\n'
                           '    9) Risks / Open Questions\n'
                           '    10) Rollout / Monitoring / Rollback\n'
                           '\n'
                           'C2. Architecture Review Checklist\n'
                           '\uf0b7     วาด user -> frontend -> API -> services -> data -> downstream\n'
                           '\uf0b7     ระบุ authentication/authorization\n'
                           '\uf0b7     ระบุ sync/async, timeout, retry, idempotency\n'
                           '\uf0b7     ระบุ source of truth และ consistency\n'
                           '\uf0b7     ระบุ sensitive-data boundary\n'
                           '\uf0b7     ระบุ failure modes + fallback\n'
                           '\uf0b7     ระบุ scaling/bottleneck/SPOF\n'
                           '\uf0b7     ระบุ logs/metrics/traces และ SLO\n'
                           '\uf0b7     ระบุ versioning/backward compatibility\n'
                           '\uf0b7     ระบุ deployment/rollback strategy\n'
                           '\n'
                           'C3. AI Production Readiness\n'
                           '\uf0b7     Evaluation dataset และ thresholds approved\n'
                           '\uf0b7     Model/prompt/retrieval config versioned\n'
                           '\uf0b7     Security/privacy review complete\n'
                           '\uf0b7     RAG access control และ data freshness verified\n'
                           '\uf0b7     Load/latency/cost test complete\n'
                           '\uf0b7     Fallback/refusal behavior tested\n'
                           '\uf0b7     Monitoring dashboard + alerts ready\n'
                           '\uf0b7     Runbook + on-call + incident owner ready\n'
                           '\uf0b7     Canary/pilot cohort defined\n'
                           '\uf0b7     Rollback/model switch tested\n'
                           '\uf0b7     Feedback capture ready\n'
                           '\uf0b7     Audit/lineage traceability verified\n'
                           '\n'
                           'C4. Incident Update Template\n'
                           ' Impact:\n'
                           ' Start time:\n'
                           ' Affected users/transactions:\n'
                           ' Current status:\n'
                           ' Suspected/confirmed root cause:\n'
                           ' Workaround/containment:\n'
                           ' Permanent fix:\n'
                           ' ETA:\n'
                           ' Monitoring/validation:\n'
                           ' Next update time:\n'
                           ' Owner:'},
               {'chapter_id': 'tpm-appendix-d',
                'chapter_title': 'Appendix D — Interview Question Bank',
                'order': 34,
                'content': 'หมวด                                                      คำถาม\n'
                           'System Design                                             Design an '
                           'internal AI knowledge assistant for 5,000\n'
                           'API                                                       What makes an API '
                           'change backward incompatible?\n'
                           'Reliability                                               Your P95 latency '
                           'jumped from 2s to 8s after release.\n'
                           'Incident                                                  A downstream '
                           'service is unavailable. How do you\n'
                           'Data                                                      Two systems show '
                           'different customer status. How do\n'
                           'AI                                                        Why would you '
                           'choose RAG instead of fine-tuning?\n'
                           'AI Evaluation                                             How would you '
                           'evaluate hallucination and\n'
                           'LLMOps                                                    How do you safely '
                           'roll out a new model version?\n'
                           'Cost                                                      Quality is 2% '
                           'better on a model that costs 4x. How\n'
                           'Platform                                                  How would you '
                           'build an internal model gateway used\n'
                           'Security                                                  How do you '
                           'protect PII in an LLM application?\n'
                           'Product                                                   A stakeholder '
                           'insists on a feature but data shows low\n'
                           'Trade-off                                                 Deadline is '
                           'fixed. Which scope do you cut and how do\n'
                           'Vendor                                                    A Chinese vendor '
                           'says the fix takes 3 weeks. How do you\n'
                           'Leadership                                                Tell me about a '
                           'technical decision where you aligned\n'
                           '\n'
                           'คำตอบสั้นที่ต้องพูดให้คล่อง\n'
                           '\n'
                           ' Why RAG instead of fine-tuning?\n'
                           ' ถ้า knowledge เปลี่ยนบ่อย RAG ทำให้ update source documents ได้โดยไม่ '
                           'retrain model ทั้งตัว และสร้าง\n'
                           ' citation/traceability ได้ง่ายกว่า Fine-tuning '
                           'เหมาะกว่าเมื่อเราต้องปรับพฤติกรรม รูปแบบ หรือ task pattern จาก\n'
                           ' examples จำนวนมาก\n'
                           '\n'
                           ' How do you manage model change in production?\n'
                           ' Version model/prompt/retrieval config -> run offline regression evaluation '
                           '-> security/quality gate ->\n'
                           ' shadow or canary -> monitor quality/latency/cost/errors -> promote or '
                           'rollback -> keep lineage so every\n'
                           ' response can be traced to a version\n'
                           '\n'
                           '30-Day Final Self-Check\n'
                           '☐ ฉันวาด Frontend -> API -> Backend -> Database -> Downstream ได้\n'
                           '\n'
                           '☐ ฉันอ่าน JSON และ HTTP status code พื้นฐานได้\n'
                           '\n'
                           '☐ ฉันอธิบาย authn vs authz ได้\n'
                           '\n'
                           '☐ ฉันอธิบาย sync/async, queue, cache, retry, idempotency ได้\n'
                           '\n'
                           '☐ ฉันเข้าใจ environment, Git, CI/CD, deployment, rollback\n'
                           '\n'
                           '☐ ฉันใช้ logs/metrics/traces และ SLO ในการคิด incident ได้\n'
                           '\n'
                           '☐ ฉันเขียน PRD/NFR/acceptance criteria ที่วัดได้\n'
                           '\n'
                           '☐ ฉันวาง test/release/migration/production readiness ได้\n'
                           '\n'
                           '☐ ฉันอธิบาย AI/ML/LLM, training/inference, token/context ได้\n'
                           '\n'
                           '☐ ฉันอธิบาย embeddings/vector search/RAG ได้\n'
                           '\n'
                           '☐ ฉันเลือก prompt vs RAG vs fine-tuning ได้\n'
                           '\n'
                           '☐ ฉันสร้าง evaluation dimensions และ hallucination controls ได้\n'
                           '\n'
                           '☐ ฉันอธิบาย AI architecture, tool calling, model gateway ได้\n'
                           '\n'
                           '☐ ฉันคำนวณ cost/request และคิด P95 latency/scale ได้\n'
                           '\n'
                           '☐ ฉันอธิบาย MLOps/LLMOps/model lifecycle/versioning ได้\n'
                           '\n'
                           '☐ ฉันเข้าใจ Platform Product และ developer experience\n'
                           '\n'
                           '☐ ฉันออกแบบ canary/shadow/A-B rollout ได้\n'
                           '\n'
                           '☐ ฉันมี AI incident/governance mindset\n'
                           '\n'
                           '☐ ฉันพูด technical Chinese ประโยคหลักได้อย่างน้อย 50 ประโยค\n'
                           '\n'
                           '☐ ฉันเล่า capstone end-to-end และตอบ trade-off ได้\n'
                           '\n'
                           '                                      END OF BOOTCAMP\n'
                           '         เป้าหมายต่อไป: ลงมือทำ Portfolio 1-2 projects และสะสม technical '
                           'delivery experience จริง'}],
  'audience': 'Mommy'},
 {'book_id': 'chinese-for-technical-product-managers',
  'title': 'Chinese for Technical Product Managers',
  'subtitle': '439 Practical Words for Systems, APIs, Core Banking, Incidents, Testing & AI',
  'author': 'Polly Study Library',
  'content_type': 'Book',
  'category': 'Chinese / Technical Product',
  'description': 'English-first Mandarin vocabulary for Technical Product Management with Chinese, '
                 'Pinyin, and word-root memory aids.',
  'cover_emoji': '🇨🇳',
  'audience': 'Mommy',
  'chapters': [{'chapter_id': 'chinese-tpm-system-api',
                'chapter_title': 'System & API',
                'order': 1,
                'content': 'System & API\n'
                           '\n'
                           'How to use this chapter\n'
                           'Read the English first and try to recall the Chinese before looking at the '
                           'next line.\n'
                           'Then check the Pinyin and use the word-root note as a memory hook.\n'
                           'The root notes are learning aids based on modern word components; they are '
                           'not claiming to be historical etymology in every case.\n'
                           '\n'
                           '1. this\n'
                           '这个\n'
                           'zhège\n'
                           'Roots: 这 = this/นี้ + 个 = general measure word/ลักษณนาม → 这个 = this '
                           'one/อันนี้\n'
                           '\n'
                           '2. be, is\n'
                           '是\n'
                           'shì\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 是 = be, is / เป็น, คือ; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '3. what\n'
                           '什么\n'
                           'shénme\n'
                           'Roots: 什 = what/อะไร (ใช้ใน 什么) + 么 = suffix in question word → 什么 = '
                           'what/อะไร\n'
                           '\n'
                           '4. Business\n'
                           '业务\n'
                           'yèwù\n'
                           'Roots: 业 = business/work/งานธุรกิจ + 务 = affairs/tasks/งาน → '
                           'business/ธุรกิจ\n'
                           '\n'
                           '5. department\n'
                           '部门\n'
                           'bùmén\n'
                           'Roots: 部 = section/ส่วน + 门 = department/gate/หน่วย → department/ฝ่าย\n'
                           '\n'
                           '6. want, think\n'
                           '想\n'
                           'xiǎng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 想 = want, think / อยาก, คิด; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '7. Resolve\n'
                           '解决\n'
                           'jiějué\n'
                           'Roots: 解 = untie/solve/แก้ + 决 = decide/resolve/ตัดสินแก้ → '
                           'resolve/แก้ปัญหา\n'
                           '\n'
                           '8. Issue\n'
                           '问题\n'
                           'wèntí\n'
                           'Roots: 问 = ask/ถาม + 题 = topic/question/ประเด็น → issue/problem/ปัญหา\n'
                           '\n'
                           '9. possessive particle\n'
                           '的\n'
                           'de\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 的 = possessive particle / ของ; '
                           'คำเชื่อมขยายคำนาม; จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '10. we\n'
                           '我们\n'
                           'wǒmen\n'
                           'Roots: 我 = I/ฉัน + 们 = plural marker/พหูพจน์ → we/พวกเรา\n'
                           '\n'
                           '11. or\n'
                           '还是\n'
                           'háishì\n'
                           'Roots: 还 = still, also / ยัง, อีก + 是 = be, is / เป็น, คือ → or / หรือ '
                           '(ในคำถามเลือก)\n'
                           '\n'
                           '12. not\n'
                           '不\n'
                           'bù\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 不 = not / ไม่; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '13. correct\n'
                           '正确\n'
                           'zhèngquè\n'
                           'Roots: 正 = right/correct/ถูก + 确 = certain/accurate/แน่นอน → '
                           'correct/ถูกต้อง\n'
                           '\n'
                           '14. completed-change particle\n'
                           '了\n'
                           'le\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 了 = completed-change particle / แล้ว; particle '
                           'แสดงการเปลี่ยนแปลง; จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '15. at, in\n'
                           '在\n'
                           'zài\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 在 = at, in / อยู่ที่, ที่; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '16. which\n'
                           '哪\n'
                           'nǎ\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 哪 = which / ไหน; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '17. one\n'
                           '一\n'
                           'yì\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 一 = one / หนึ่ง; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '18. after\n'
                           '后\n'
                           'hòu\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 后 = after / หลัง; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '19. not have\n'
                           '没有\n'
                           'méiyǒu\n'
                           'Roots: 没 = not have/ไม่มี + 有 = have/มี → not have/ไม่มี\n'
                           '\n'
                           '20. call, invoke\n'
                           '调用\n'
                           'diàoyòng\n'
                           'Roots: 调 = call/dispatch/เรียก + 用 = use/ใช้ → invoke/call/เรียกใช้งาน\n'
                           '\n'
                           '21. pass, transmit\n'
                           '传\n'
                           'chuán\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 传 = pass, transmit / ส่ง; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '22. code\n'
                           '码\n'
                           'mǎ\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 码 = code / รหัส; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '23. come from\n'
                           '来自\n'
                           'láizì\n'
                           'Roots: 来 = come/มา + 自 = from/self/จาก → come from/มาจาก\n'
                           '\n'
                           '24. which\n'
                           '哪个\n'
                           'nǎge\n'
                           'Roots: 哪 = which / ไหน + 个 = general measure word / ลักษณนามทั่วไป → which '
                           '/ อันไหน, ระบบไหน\n'
                           '\n'
                           '25. please\n'
                           '请\n'
                           'qǐng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 请 = please / กรุณา; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '26. check, inspect\n'
                           '检查\n'
                           'jiǎnchá\n'
                           'Roots: 检 = inspect/ตรวจ + 查 = check/search/เช็ก → inspect/check/ตรวจสอบ\n'
                           '\n'
                           '27. inside\n'
                           '里\n'
                           'lǐ\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 里 = inside / ใน, ข้างใน; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '28. can\n'
                           '可以\n'
                           'kěyǐ\n'
                           'Roots: 可 = can/ได้ + 以 = by/able to/สามารถ → can/สามารถ\n'
                           '\n'
                           '29. be, become\n'
                           '为\n'
                           'wéi\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 为 = be, become / เป็น, กลายเป็น; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '30. empty, null\n'
                           '空\n'
                           'kōng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 空 = empty, null / ว่าง; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '31. question particle\n'
                           '吗\n'
                           'ma\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 吗 = question particle / ไหม; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '32. two\n'
                           '两\n'
                           'liǎng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 两 = two / สอง; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '33. general measure word\n'
                           '个\n'
                           'ge\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 个 = general measure word / ลักษณนามทั่วไป; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '34. consistent\n'
                           '一致\n'
                           'yízhì\n'
                           'Roots: 一 = one/same/หนึ่ง + 致 = reach/consistent/ตรงกัน → '
                           'consistent/สอดคล้อง\n'
                           '\n'
                           '35. give, for\n'
                           '给\n'
                           'gěi\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 给 = give, for / ให้; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '36. I\n'
                           '我\n'
                           'wǒ\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 我 = I / ฉัน, ผม; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '37. a bit\n'
                           '一下\n'
                           'yíxià\n'
                           'Roots: 一 = one / หนึ่ง + 下 = down/off / ลง → a bit / สักหน่อย\n'
                           '\n'
                           '38. from\n'
                           '从\n'
                           'cóng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 从 = from / จาก; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '39. to\n'
                           '到\n'
                           'dào\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 到 = to / ถึง, ไปยัง; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '40. if\n'
                           '如果\n'
                           'rúguǒ\n'
                           'Roots: 如 = as/if/หาก + 果 = result/ผล → if/ถ้า (จำทั้งคำ)\n'
                           '\n'
                           '41. available\n'
                           '可用\n'
                           'kěyòng\n'
                           'Roots: 可 = can/สามารถ + 用 = use/ใช้ → available/usable/พร้อมใช้\n'
                           '\n'
                           '42. have\n'
                           '有\n'
                           'yǒu\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 有 = have / มี; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '43. now\n'
                           '现在\n'
                           'xiànzài\n'
                           'Roots: 现 = present/ปรากฏ + 在 = be at/อยู่ → now/ตอนนี้\n'
                           '\n'
                           '44. time, when\n'
                           '时候\n'
                           'shíhou\n'
                           'Roots: 时 = time/เวลา + 候 = time/occasion/ช่วง → time/ตอนที่\n'
                           '\n'
                           '45. how much, how many\n'
                           '多少\n'
                           'duōshao\n'
                           'Roots: 多 = many/much/มาก + 少 = few/little/น้อย → how much/how many/เท่าไร'},
               {'chapter_id': 'chinese-tpm-incident-operations',
                'chapter_title': 'Incident & Operations',
                'order': 2,
                'content': 'Incident & Operations\n'
                           '\n'
                           'How to use this chapter\n'
                           'Read the English first and try to recall the Chinese before looking at the '
                           'next line.\n'
                           'Then check the Pinyin and use the word-root note as a memory hook.\n'
                           'The root notes are learning aids based on modern word components; they are '
                           'not claiming to be historical etymology in every case.\n'
                           '\n'
                           '1. customer\n'
                           '客户\n'
                           'kèhù\n'
                           'Roots: 客 = guest/client/ลูกค้า + 户 = household/account/ผู้ใช้ → '
                           'customer/ลูกค้า\n'
                           '\n'
                           '2. time\n'
                           '时间\n'
                           'shíjiān\n'
                           'Roots: 时 = time/เวลา + 间 = interval/ช่วง → time/เวลา\n'
                           '\n'
                           '3. obvious, significant\n'
                           '明显\n'
                           'míngxiǎn\n'
                           'Roots: 明 = clear/ชัด + 显 = visible/เด่น → obvious/significant/ชัดเจน\n'
                           '\n'
                           '4. Incident\n'
                           '事故\n'
                           'shìgù\n'
                           'Roots: 事 = event/matter/เหตุการณ์ + 故 = cause/incident/เหตุ → '
                           'incident/เหตุขัดข้อง\n'
                           '\n'
                           '5. do\n'
                           '做\n'
                           'zuò\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 做 = do / ทำ; จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '6. Outcome\n'
                           '结果\n'
                           'jiéguǒ\n'
                           'Roots: 结 = tie/form/เกิดผล + 果 = fruit/result/ผล → outcome/result/ผลลัพธ์\n'
                           '\n'
                           '7. merely, only\n'
                           '仅仅\n'
                           'jǐnjǐn\n'
                           'Roots: 仅 = only/เพียง + 仅 = only/ย้ำความหมาย → merely/เพียงแค่\n'
                           '\n'
                           '8. most\n'
                           '最\n'
                           'zuì\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 最 = most / ที่สุด; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '9. big\n'
                           '大\n'
                           'dà\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 大 = big / ใหญ่; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '10. who\n'
                           '谁\n'
                           'shéi\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 谁 = who / ใคร; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '11. still, also\n'
                           '还\n'
                           'hái\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 还 = still, also / ยัง, อีก; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '12. enough\n'
                           '够\n'
                           'gòu\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 够 = enough / พอ; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '13. non-\n'
                           '非\n'
                           'fēi\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 非 = non- / ไม่ใช่, non-; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '14. functional\n'
                           '功能性\n'
                           'gōngnéngxìng\n'
                           'Roots: 功能 = Function / ฟังก์ชัน + 性 = property/-ness / คุณสมบัติ/ภาวะ → '
                           'functional / เชิงฟังก์ชัน\n'
                           '\n'
                           '15. must\n'
                           '必须\n'
                           'bìxū\n'
                           'Roots: 必 = must/ต้อง + 须 = must/จำเป็น → must/ต้อง\n'
                           '\n'
                           '16. exceed\n'
                           '超过\n'
                           'chāoguò\n'
                           'Roots: 超 = exceed/เกิน + 过 = pass/ผ่าน → exceed/เกิน\n'
                           '\n'
                           '17. second\n'
                           '秒\n'
                           'miǎo\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 秒 = second / วินาที; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '18. this\n'
                           '这\n'
                           'zhè\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 这 = this / นี้; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '19. kind, type\n'
                           '种\n'
                           'zhǒng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 种 = kind, type / ชนิด; ลักษณนาม; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '20. situation\n'
                           '情况\n'
                           'qíngkuàng\n'
                           'Roots: 情 = condition/situation/สภาพ + 况 = condition/state/ภาวะ → '
                           'situation/สถานการณ์\n'
                           '\n'
                           '21. under, below\n'
                           '下\n'
                           'xià\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 下 = under, below / ภายใต้, ด้านล่าง; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '22. should\n'
                           '应该\n'
                           'yīnggāi\n'
                           'Roots: 应 = should/respond/ควร + 该 = should/ควร → should/ควร\n'
                           '\n'
                           '23. Handle\n'
                           '处理\n'
                           'chǔlǐ\n'
                           'Roots: 处 = handle/deal with/จัดการ + 理 = manage/reason/จัดระเบียบ → '
                           'handle/จัดการ\n'
                           '\n'
                           '24. Acceptance\n'
                           '验收\n'
                           'yànshōu\n'
                           'Roots: 验 = inspect/test/ตรวจ + 收 = receive/รับ → acceptance/ตรวจแล้วรับ\n'
                           '\n'
                           '25. Mandatory\n'
                           '必填\n'
                           'bìtián\n'
                           'Roots: 必 = must/ต้อง + 填 = fill/กรอก → mandatory/ต้องกรอก\n'
                           '\n'
                           '26. Optional\n'
                           '可选\n'
                           'kěxuǎn\n'
                           'Roots: 可 = can/สามารถ + 选 = choose/เลือก → optional/เลือกได้\n'
                           '\n'
                           '27. Change\n'
                           '变更\n'
                           'biàngēng\n'
                           'Roots: 变 = change/เปลี่ยน + 更 = alter/update/ปรับ → change/การเปลี่ยนแปลง\n'
                           '\n'
                           '28. backward\n'
                           '向后\n'
                           'xiànghòu\n'
                           'Roots: 向 = toward/ไปทาง + 后 = back/หลัง → backward/ไปทางด้านหลัง\n'
                           '\n'
                           '29. Compatible\n'
                           '兼容\n'
                           'jiānróng\n'
                           'Roots: 兼 = combine/together/ร่วมกัน + 容 = contain/accommodate/รองรับ → '
                           'compatible/เข้ากันได้\n'
                           '\n'
                           '30. supplement, add\n'
                           '补充\n'
                           'bǔchōng\n'
                           'Roots: 补 = supplement/fill/เติม + 充 = fill/เติมเต็ม → '
                           'supplement/add/เพิ่มเติม\n'
                           '\n'
                           '31. Scenario\n'
                           '场景\n'
                           'chǎngjǐng\n'
                           'Roots: 场 = place/scene/สถานที่ + 景 = scene/view/ภาพ → scenario/use '
                           'scene/สถานการณ์\n'
                           '\n'
                           '32. task\n'
                           '任务\n'
                           'rènwù\n'
                           'Roots: 任 = assign/รับหน้าที่ + 务 = task/งาน → task/งานที่รับผิดชอบ\n'
                           '\n'
                           '33. Complete\n'
                           '完成\n'
                           'wánchéng\n'
                           'Roots: 完 = complete/ครบ + 成 = become/finish/สำเร็จ → complete/เสร็จ\n'
                           '\n'
                           '34. discover, find\n'
                           '发现\n'
                           'fāxiàn\n'
                           'Roots: 发 = bring out/ทำให้เกิด + 现 = appear/ปรากฏ → discover/find/พบ '
                           '(จำความหมายรวม)\n'
                           '\n'
                           '35. Defect\n'
                           '缺陷\n'
                           'quēxiàn\n'
                           'Roots: 缺 = lack/missing/ขาด + 陷 = defect/pit/จุดบกพร่อง → '
                           'defect/ข้อบกพร่อง\n'
                           '\n'
                           '36. will, can\n'
                           '会\n'
                           'huì\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 会 = will, can / จะ, สามารถ; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '37. Regression\n'
                           '回归\n'
                           'huíguī\n'
                           'Roots: 回 = return/กลับ + 归 = return to/กลับคืน → '
                           'regression/กลับไปตรวจของเดิม\n'
                           '\n'
                           '38. Performance\n'
                           '性能\n'
                           'xìngnéng\n'
                           'Roots: 性 = property/คุณสมบัติ + 能 = ability/capability/ความสามารถ → '
                           'performance/ประสิทธิภาพ\n'
                           '\n'
                           '39. how is it\n'
                           '怎么样\n'
                           'zěnmeyàng\n'
                           'Roots: 怎么 = how/อย่างไร + 样 = manner/type/แบบ → how is it?/เป็นอย่างไร\n'
                           '\n'
                           '40. personal\n'
                           '个人\n'
                           'gèrén\n'
                           'Roots: 个 = individual/หนึ่งราย + 人 = person/คน → personal/ส่วนบุคคล\n'
                           '\n'
                           '41. information\n'
                           '信息\n'
                           'xìnxī\n'
                           'Roots: 信 = message/information/สาร + 息 = information/news/ข่าวสาร → '
                           'information/ข้อมูล\n'
                           '\n'
                           '42. Dependency\n'
                           '依赖\n'
                           'yīlài\n'
                           'Roots: 依 = rely/พึ่ง + 赖 = depend on/อาศัย → dependency/การพึ่งพา\n'
                           '\n'
                           '43. approximately\n'
                           '大概\n'
                           'dàgài\n'
                           'Roots: 大 = big/general/กว้าง ๆ + 概 = general idea/โดยประมาณ → '
                           'approximately/ประมาณ\n'
                           '\n'
                           '44. many, much\n'
                           '多\n'
                           'duō\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 多 = many, much / มาก; เท่าไร; จำเป็นหน่วยคำตรง '
                           'ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '45. long\n'
                           '长\n'
                           'cháng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 长 = long / ยาว, นาน; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '46. reduce, shrink\n'
                           '缩小\n'
                           'suōxiǎo\n'
                           'Roots: 缩 = shrink/หด + 小 = small/เล็ก → reduce/shrink/ลด\n'
                           '\n'
                           '47. technology, technical\n'
                           '技术\n'
                           'jìshù\n'
                           'Roots: 技 = skill/เทคนิค + 术 = method/art/วิธี → '
                           'technology/technical/เทคนิค\n'
                           '\n'
                           '48. debt\n'
                           '债\n'
                           'zhài\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 债 = debt / หนี้; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '49. accept\n'
                           '接受\n'
                           'jiēshòu\n'
                           'Roots: 接 = receive/รับ + 受 = accept/รับไว้ → accept/ยอมรับ\n'
                           '\n'
                           '50. these\n'
                           '这些\n'
                           'zhèxiē\n'
                           'Roots: 这 = this/นี้ + 些 = some/plural/หลายรายการ → these/เหล่านี้\n'
                           '\n'
                           '51. contain, include\n'
                           '包含\n'
                           'bāohán\n'
                           'Roots: 包 = wrap/include/ห่อรวม + 含 = contain/มีอยู่ → '
                           'contain/include/ประกอบด้วย\n'
                           '\n'
                           '52. Permission\n'
                           '权限\n'
                           'quánxiàn\n'
                           'Roots: 权 = authority/right/สิทธิ์ + 限 = limit/ขอบเขต → '
                           'permission/สิทธิ์ที่มีขอบเขต\n'
                           '\n'
                           '53. access\n'
                           '访问\n'
                           'fǎngwèn\n'
                           'Roots: 访 = visit/เข้าหา + 问 = ask/ถาม → access/visit/เข้าถึง\n'
                           '\n'
                           '54. whether\n'
                           '是否\n'
                           'shìfǒu\n'
                           'Roots: 是 = is/คือ + 否 = not/no/ไม่ → whether or not/หรือไม่\n'
                           '\n'
                           '55. already\n'
                           '已经\n'
                           'yǐjīng\n'
                           'Roots: 已 = already/แล้ว + 经 = pass/already/ผ่านแล้ว → '
                           'already/เรียบร้อยแล้ว\n'
                           '\n'
                           '56. Encryption\n'
                           '加密\n'
                           'jiāmì\n'
                           'Roots: 加 = add/เพิ่ม + 密 = secret/dense/ลับ → '
                           'encryption/เพิ่มความลับให้ข้อมูล\n'
                           '\n'
                           '57. record\n'
                           '记录\n'
                           'jìlù\n'
                           'Roots: 记 = remember/record/บันทึก + 录 = record/list/ลงรายการ → '
                           'record/บันทึก\n'
                           '\n'
                           '58. Audit\n'
                           '审计\n'
                           'shěnjì\n'
                           'Roots: 审 = examine/review/ตรวจ + 计 = calculate/record/บันทึก → '
                           'audit/ตรวจสอบย้อนหลัง\n'
                           '\n'
                           '59. Service\n'
                           '服务\n'
                           'fúwù\n'
                           'Roots: 服 = serve/รับใช้ + 务 = duty/task/งาน → service/บริการ\n'
                           '\n'
                           '60. account\n'
                           '账号\n'
                           'zhànghào\n'
                           'Roots: 账 = account/บัญชี + 号 = number/id/หมายเลข → account/บัญชีผู้ใช้\n'
                           '\n'
                           '61. Deployment\n'
                           '部署\n'
                           'bùshǔ\n'
                           'Roots: 部 = part/section/ส่วน + 署 = arrange/deploy/จัดวาง → '
                           'deployment/นำระบบไปวาง\n'
                           '\n'
                           '62. plan\n'
                           '计划\n'
                           'jìhuà\n'
                           'Roots: 计 = plan/calculate/วางแผน + 划 = plan/divide/กำหนด → plan/แผน\n'
                           '\n'
                           '63. Migration\n'
                           '迁移\n'
                           'qiānyí\n'
                           'Roots: 迁 = move/ย้าย + 移 = shift/เคลื่อน → migration/การย้าย\n'
                           '\n'
                           '64. Reconciliation\n'
                           '核对\n'
                           'héduì\n'
                           'Roots: 核 = check/core/ตรวจหลัก + 对 = compare/match/เทียบ → '
                           'reconciliation/ตรวจเทียบ\n'
                           '\n'
                           '65. need to, will\n'
                           '要\n'
                           'yào\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 要 = need to, will / ต้อง, จะ; จำเป็นหน่วยคำตรง '
                           'ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '66. Monitoring\n'
                           '监控\n'
                           'jiānkòng\n'
                           'Roots: 监 = watch/เฝ้าดู + 控 = control/ควบคุม → monitoring/เฝ้าระวัง\n'
                           '\n'
                           '67. Metric\n'
                           '指标\n'
                           'zhǐbiāo\n'
                           'Roots: 指 = point/ชี้ + 标 = mark/standard/เครื่องหมาย → metric/ตัวชี้วัด\n'
                           '\n'
                           '68. tonight\n'
                           '今晚\n'
                           'jīnwǎn\n'
                           'Roots: 今 = current/today/วันนี้ + 晚 = evening/night/กลางคืน → '
                           'tonight/คืนนี้\n'
                           '\n'
                           '69. production\n'
                           '生产\n'
                           'shēngchǎn\n'
                           'Roots: 生 = produce/born/เกิด + 产 = produce/ผลิต → production/การผลิต\n'
                           '\n'
                           '70. machine\n'
                           '机器\n'
                           'jīqì\n'
                           'Roots: 机 = machine/mechanism/เครื่อง + 器 = device/อุปกรณ์ → '
                           'machine/เครื่องจักร\n'
                           '\n'
                           '71. learn, learning\n'
                           '学习\n'
                           'xuéxí\n'
                           'Roots: 学 = study/เรียน + 习 = practice/ฝึก → learning/เรียนรู้\n'
                           '\n'
                           '72. language\n'
                           '语言\n'
                           'yǔyán\n'
                           'Roots: 语 = language/speech/ภาษา + 言 = speech/คำพูด → language/ภาษา\n'
                           '\n'
                           '73. Model\n'
                           '模型\n'
                           'móxíng\n'
                           'Roots: 模 = model/pattern/แบบ + 型 = type/form/รูปแบบ → model/โมเดล\n'
                           '\n'
                           '74. application\n'
                           '应用\n'
                           'yìngyòng\n'
                           'Roots: 应 = apply/respond/นำไปใช้ + 用 = use/ใช้ → application/use/การใช้งาน\n'
                           '\n'
                           '75. all\n'
                           '所有\n'
                           'suǒyǒu\n'
                           'Roots: 所 = all/that which/สิ่งที่ + 有 = have/มี → all/ทั้งหมด\n'
                           '\n'
                           '76. all\n'
                           '都\n'
                           'dōu\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 都 = all / ทั้งหมด, ล้วน; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '77. manual, human\n'
                           '人工\n'
                           'réngōng\n'
                           'Roots: 人 = human/คน + 工 = work/made/ทำ → human-made/manual/โดยคน\n'
                           '\n'
                           '78. intelligence\n'
                           '智能\n'
                           'zhìnéng\n'
                           'Roots: 智 = wisdom/intelligence/ปัญญา + 能 = ability/ความสามารถ → '
                           'intelligence/ความฉลาด\n'
                           '\n'
                           '79. only, merely\n'
                           '只是\n'
                           'zhǐshì\n'
                           'Roots: 只 = only/เพียง + 是 = is/คือ → only/เพียงแค่\n'
                           '\n'
                           '80. part\n'
                           '部分\n'
                           'bùfen\n'
                           'Roots: 部 = part/ส่วน + 分 = divide/ส่วนแบ่ง → part/ส่วนหนึ่ง\n'
                           '\n'
                           '81. Output\n'
                           '输出\n'
                           'shūchū\n'
                           'Roots: 输 = transmit/output/ส่ง + 出 = out/ออก → output/ผลที่ส่งออก\n'
                           '\n'
                           '82. exist\n'
                           '存在\n'
                           'cúnzài\n'
                           'Roots: 存 = exist/store/มีอยู่ + 在 = be at/อยู่ → exist/มีอยู่\n'
                           '\n'
                           '83. certainty\n'
                           '确定性\n'
                           'quèdìngxìng\n'
                           'Roots: 确定 = certain/แน่นอน + 性 = property/state/ภาวะ → '
                           'certainty/ความแน่นอน\n'
                           '\n'
                           '84. define\n'
                           '定义\n'
                           'dìngyì\n'
                           'Roots: 定 = set/fix/กำหนด + 义 = meaning/ความหมาย → define/นิยาม\n'
                           '\n'
                           '85. tolerance\n'
                           '容忍度\n'
                           'róngrěndù\n'
                           'Roots: 容忍 = tolerate/ยอมรับ + 度 = degree/ระดับ → tolerance/ระดับการยอมรับ\n'
                           '\n'
                           '86. Training\n'
                           '训练\n'
                           'xùnliàn\n'
                           'Roots: 训 = train/instruct/ฝึก + 练 = practice/ฝึกซ้ำ → training/การฝึก\n'
                           '\n'
                           '87. Inference\n'
                           '推理\n'
                           'tuīlǐ\n'
                           'Roots: 推 = infer/push/อนุมาน + 理 = reason/เหตุผล → inference/การอนุมาน'},
               {'chapter_id': 'chinese-tpm-cbs-core-banking',
                'chapter_title': 'CBS · Core Banking System',
                'order': 3,
                'content': 'CBS · Core Banking System\n'
                           '\n'
                           'How to use this chapter\n'
                           'Read the English first and try to recall the Chinese before looking at the '
                           'next line.\n'
                           'Then check the Pinyin and use the word-root note as a memory hook.\n'
                           'The root notes are learning aids based on modern word components; they are '
                           'not claiming to be historical etymology in every case.\n'
                           '\n'
                           '1. Requirement\n'
                           '需求\n'
                           'xūqiú\n'
                           'Roots: 需 = need/ต้องการ + 求 = seek/request/แสวงหา → '
                           'requirement/ความต้องการ\n'
                           '\n'
                           '2. Project\n'
                           '项目\n'
                           'xiàngmù\n'
                           'Roots: 项 = item/รายการ + 目 = goal/item/หัวข้อ → project/โครงการ\n'
                           '\n'
                           '3. Scope\n'
                           '范围\n'
                           'fànwéi\n'
                           'Roots: 范 = scope/pattern/ขอบเขต + 围 = surround/ล้อม → ขอบเขตที่ล้อมไว้\n'
                           '\n'
                           '4. need\n'
                           '需要\n'
                           'xūyào\n'
                           'Roots: 需 = need/ต้องการ + 要 = need/want/ต้อง → need/จำเป็นต้อง\n'
                           '\n'
                           '5. first, beforehand\n'
                           '先\n'
                           'xiān\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 先 = first, beforehand / ก่อน; จำเป็นหน่วยคำตรง '
                           'ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '6. Confirm\n'
                           '确认\n'
                           'quèrèn\n'
                           'Roots: 确 = certain/correct/แน่นอน + 认 = recognize/confirm/ยอมรับ → '
                           'confirm/ยืนยัน\n'
                           '\n'
                           '7. Product\n'
                           '产品\n'
                           'chǎnpǐn\n'
                           'Roots: 产 = produce/ผลิต + 品 = product/item/สินค้า → product/ผลิตภัณฑ์\n'
                           '\n'
                           '8. Goal\n'
                           '目标\n'
                           'mùbiāo\n'
                           'Roots: 目 = eye/goal/เป้า + 标 = mark/target/เครื่องหมาย → goal/เป้าหมาย\n'
                           '\n'
                           '9. Frontend\n'
                           '前端\n'
                           'qiánduān\n'
                           'Roots: 前 = front/หน้า + 端 = end/side/ด้าน → frontend/ฝั่งหน้า\n'
                           '\n'
                           '10. Backend\n'
                           '后端\n'
                           'hòuduān\n'
                           'Roots: 后 = back/หลัง + 端 = end/side/ด้าน → backend/ฝั่งหลัง\n'
                           '\n'
                           '11. display\n'
                           '显示\n'
                           'xiǎnshì\n'
                           'Roots: 显 = make visible/ชัด + 示 = show/แสดง → display/แสดง\n'
                           '\n'
                           '12. Data\n'
                           '数据\n'
                           'shùjù\n'
                           'Roots: 数 = number/count/จำนวน + 据 = data/evidence/ข้อมูลอ้างอิง → '
                           'data/ข้อมูล\n'
                           '\n'
                           '13. Return\n'
                           '返回\n'
                           'fǎnhuí\n'
                           'Roots: 返 = return/ย้อนกลับ + 回 = return/back/กลับ → return/คืนค่า\n'
                           '\n'
                           '14. Occur\n'
                           '发生\n'
                           'fāshēng\n'
                           'Roots: 发 = occur/send out/เกิด + 生 = arise/be born/เกิดขึ้น → '
                           'occur/เกิดขึ้น\n'
                           '\n'
                           '15. layer\n'
                           '层\n'
                           'céng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 层 = layer / ชั้น; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '16. User\n'
                           '用户\n'
                           'yònghù\n'
                           'Roots: 用 = use/ใช้ + 户 = household/user/ผู้ใช้ → user/ผู้ใช้งาน\n'
                           '\n'
                           '17. click\n'
                           '点击\n'
                           'diǎnjī\n'
                           'Roots: 点 = point/tap/จุด + 击 = strike/click/กด → click/tap/กด\n'
                           '\n'
                           '18. button\n'
                           '按钮\n'
                           'ànniǔ\n'
                           'Roots: 按 = press/กด + 钮 = button/ปุ่ม → button/ปุ่ม\n'
                           '\n'
                           '19. Response\n'
                           '响应\n'
                           'xiǎngyìng\n'
                           'Roots: 响 = sound/echo/เสียงตอบ + 应 = respond/ตอบสนอง → response/การตอบกลับ\n'
                           '\n'
                           '20. how\n'
                           '怎么\n'
                           'zěnme\n'
                           'Roots: 怎 = how/อย่างไร + 么 = question suffix → how/อย่างไร\n'
                           '\n'
                           '21. Request\n'
                           '请求\n'
                           'qǐngqiú\n'
                           'Roots: 请 = request/ขอ + 求 = seek/request/ร้องขอ → request/คำขอ\n'
                           '\n'
                           '22. which ones\n'
                           '哪些\n'
                           'nǎxiē\n'
                           'Roots: 哪 = which/ไหน + 些 = some/plural/บางรายการ → which ones/อะไรบ้าง\n'
                           '\n'
                           '23. Parameter\n'
                           '参数\n'
                           'cānshù\n'
                           'Roots: 参 = participate/reference/อ้างอิง + 数 = number/value/ค่า → '
                           'parameter/พารามิเตอร์\n'
                           '\n'
                           '24. use\n'
                           '使用\n'
                           'shǐyòng\n'
                           'Roots: 使 = make/use/ใช้ + 用 = use/ใช้ → use/ใช้งาน\n'
                           '\n'
                           '25. Token\n'
                           '令牌\n'
                           'lìngpái\n'
                           'Roots: 令 = command/authorization/คำสั่ง + 牌 = badge/token/ป้าย → '
                           'token/โทเคน\n'
                           '\n'
                           '26. Error\n'
                           '错误\n'
                           'cuòwù\n'
                           'Roots: 错 = wrong/ผิด + 误 = mistake/error/พลาด → error/ข้อผิดพลาด\n'
                           '\n'
                           '27. when\n'
                           '时\n'
                           'shí\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 时 = when / เมื่อ, ตอนที่; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '28. System\n'
                           '系统\n'
                           'xìtǒng\n'
                           'Roots: 系 = connect/link/เชื่อมโยง + 统 = unify/รวมเป็นหนึ่ง → system/ระบบ\n'
                           '\n'
                           '29. Database\n'
                           '数据库\n'
                           'shùjùkù\n'
                           'Roots: 数据 = data/ข้อมูล + 库 = repository/store/คลัง → database/ฐานข้อมูล\n'
                           '\n'
                           '30. main, primary\n'
                           '主要\n'
                           'zhǔyào\n'
                           'Roots: 主 = main/หลัก + 要 = important/สำคัญ → main/primary/หลัก\n'
                           '\n'
                           '31. Data Source\n'
                           '数据源\n'
                           'shùjùyuán\n'
                           'Roots: 数据 = data / ข้อมูล + 源 = source / ต้นทาง/แหล่ง → Data Source / '
                           'แหล่งข้อมูล\n'
                           '\n'
                           '32. Field\n'
                           '字段\n'
                           'zìduàn\n'
                           'Roots: 字 = character/ตัวอักษร + 段 = section/segment/ส่วน → field/ฟิลด์\n'
                           '\n'
                           '33. explain\n'
                           '解释\n'
                           'jiěshì\n'
                           'Roots: 解 = explain/solve/แก้ความ + 释 = explain/release/คลี่ความ → '
                           'explain/อธิบาย\n'
                           '\n'
                           '34. Architecture\n'
                           '架构\n'
                           'jiàgòu\n'
                           'Roots: 架 = frame/support/โครง + 构 = construct/build/ประกอบ → '
                           'architecture/สถาปัตยกรรม\n'
                           '\n'
                           '35. Connect\n'
                           '连接\n'
                           'liánjiē\n'
                           'Roots: 连 = connect/ต่อ + 接 = join/receive/เชื่อม → connect/เชื่อมต่อ\n'
                           '\n'
                           '36. step\n'
                           '步骤\n'
                           'bùzhòu\n'
                           'Roots: 步 = step/ก้าว + 骤 = sequence/stage/ช่วง → step/ขั้นตอน\n'
                           '\n'
                           '37. Synchronous\n'
                           '同步\n'
                           'tóngbù\n'
                           'Roots: 同 = same/พร้อมกัน + 步 = step/ก้าว → synchronous/ก้าวไปพร้อมกัน\n'
                           '\n'
                           '38. Asynchronous\n'
                           '异步\n'
                           'yìbù\n'
                           'Roots: 异 = different/ต่าง + 步 = step/ก้าว → asynchronous/ก้าวไม่พร้อมกัน\n'
                           '\n'
                           '39. backup, standby\n'
                           '备用\n'
                           'bèiyòng\n'
                           'Roots: 备 = prepare/backup/สำรอง + 用 = use/ใช้ → backup/สำรอง\n'
                           '\n'
                           '40. plan, solution\n'
                           '方案\n'
                           'fāng’àn\n'
                           'Roots: 方 = method/direction/แนวทาง + 案 = plan/case/แผน → '
                           'solution/plan/แนวทาง\n'
                           '\n'
                           '41. Environment\n'
                           '环境\n'
                           'huánjìng\n'
                           'Roots: 环 = surround/ล้อม + 境 = environment/boundary/สภาพ → '
                           'environment/สภาพแวดล้อม\n'
                           '\n'
                           '42. Version\n'
                           '版本\n'
                           'bǎnběn\n'
                           'Roots: 版 = edition/version/ฉบับ + 本 = base/copy/ฉบับต้น → version/เวอร์ชัน\n'
                           '\n'
                           '43. Go-live\n'
                           '上线\n'
                           'shàngxiàn\n'
                           'Roots: 上 = up/go on/ขึ้น + 线 = line/network/เส้นระบบ → go-live/ขึ้นระบบ\n'
                           '\n'
                           '44. Pipeline\n'
                           '流水线\n'
                           'liúshuǐxiàn\n'
                           'Roots: 流水 = flowing water/การไหล + 线 = line/เส้น → '
                           'pipeline/งานไหลเป็นลำดับ\n'
                           '\n'
                           '45. fail\n'
                           '失败\n'
                           'shībài\n'
                           'Roots: 失 = lose/fail/เสีย + 败 = defeat/fail/พ่าย → fail/ล้มเหลว\n'
                           '\n'
                           '46. appear, occur\n'
                           '出现\n'
                           'chūxiàn\n'
                           'Roots: 出 = out/ออก + 现 = appear/ปรากฏ → appear/เกิดขึ้น\n'
                           '\n'
                           '47. Rollback\n'
                           '回滚\n'
                           'huígǔn\n'
                           'Roots: 回 = return/back/กลับ + 滚 = roll/กลิ้งย้อน → rollback/ย้อนเวอร์ชัน\n'
                           '\n'
                           '48. carry out, proceed\n'
                           '进行\n'
                           'jìnxíng\n'
                           'Roots: 进 = advance/เดินหน้า + 行 = carry out/ดำเนิน → proceed/carry '
                           'out/ดำเนินการ\n'
                           '\n'
                           '49. canary\n'
                           '金丝雀\n'
                           'jīnsīquè\n'
                           'Roots: 金丝雀 = canary bird/นกคานารี; ใน tech ใช้เป็นภาพจำของ Canary release '
                           'คือปล่อยให้กลุ่มเล็กก่อนแล้วค่อยขยาย\n'
                           '\n'
                           '50. Release\n'
                           '发布\n'
                           'fābù\n'
                           'Roots: 发 = send/release/ปล่อย + 布 = spread/กระจาย → '
                           'release/publish/ปล่อยเวอร์ชัน\n'
                           '\n'
                           '51. provide\n'
                           '提供\n'
                           'tígōng\n'
                           'Roots: 提 = raise/provide/ยกให้ + 供 = supply/จัดหา → provide/จัดให้\n'
                           '\n'
                           '52. Log\n'
                           '日志\n'
                           'rìzhì\n'
                           'Roots: 日 = day/วัน + 志 = record/log/บันทึก → log/บันทึกระบบ\n'
                           '\n'
                           '53. fundamental, root\n'
                           '根本\n'
                           'gēnběn\n'
                           'Roots: 根 = root/ราก + 本 = base/origin/ต้น → root/fundamental/รากฐาน\n'
                           '\n'
                           '54. reason, cause\n'
                           '原因\n'
                           'yuányīn\n'
                           'Roots: 原 = origin/ต้นกำเนิด + 因 = cause/เหตุ → cause/reason/สาเหตุ\n'
                           '\n'
                           '55. Impact\n'
                           '影响\n'
                           'yǐngxiǎng\n'
                           'Roots: 影 = shadow/effect/เงา + 响 = echo/respond/สะท้อน → impact/ผลกระทบ\n'
                           '\n'
                           '56. increase\n'
                           '增加\n'
                           'zēngjiā\n'
                           'Roots: 增 = increase/เพิ่ม + 加 = add/บวกเพิ่ม → increase/เพิ่มขึ้น\n'
                           '\n'
                           '57. Postmortem\n'
                           '复盘\n'
                           'fùpán\n'
                           'Roots: 复 = repeat/review again/ทบทวนอีกครั้ง + 盘 = board/overall '
                           'picture/กระดาน → postmortem/review\n'
                           '\n'
                           '58. measure\n'
                           '衡量\n'
                           'héngliáng\n'
                           'Roots: 衡 = weigh/measure/ชั่ง + 量 = measure/quantity/วัด → measure/วัดผล\n'
                           '\n'
                           '59. while, but\n'
                           '而\n'
                           'ér\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 而 = while, but / และ/แต่; คำเชื่อม; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '60. output\n'
                           '产出\n'
                           'chǎnchū\n'
                           'Roots: 产 = produce/ผลิต + 出 = out/ออก → output/ผลผลิต\n'
                           '\n'
                           '61. Risk\n'
                           '风险\n'
                           'fēngxiǎn\n'
                           'Roots: 风 = wind/change/ลม + 险 = danger/risk/อันตราย → risk/ความเสี่ยง\n'
                           '\n'
                           '62. Assumption\n'
                           '假设\n'
                           'jiǎshè\n'
                           'Roots: 假 = suppose/สมมติ + 设 = set/ตั้ง → assumption/สมมติฐาน\n'
                           '\n'
                           '63. Validate\n'
                           '验证\n'
                           'yànzhèng\n'
                           'Roots: 验 = test/check/ตรวจ + 证 = prove/evidence/พิสูจน์ → '
                           'validate/ตรวจยืนยัน\n'
                           '\n'
                           '64. Clarify\n'
                           '明确\n'
                           'míngquè\n'
                           'Roots: 明 = clear/ชัด + 确 = certain/แน่นอน → clarify/make clear/ทำให้ชัด\n'
                           '\n'
                           '65. Exception\n'
                           '异常\n'
                           'yìcháng\n'
                           'Roots: 异 = different/ผิดปกติ + 常 = normal/ปกติ → exception/ความผิดปกติ\n'
                           '\n'
                           '66. standard\n'
                           '标准\n'
                           'biāozhǔn\n'
                           'Roots: 标 = mark/standard/เกณฑ์ + 准 = accurate/standard/มาตรฐาน → '
                           'standard/มาตรฐาน\n'
                           '\n'
                           '67. Testing\n'
                           '测试\n'
                           'cèshì\n'
                           'Roots: 测 = measure/test/วัดทดสอบ + 试 = try/test/ทดลอง → testing/การทดสอบ\n'
                           '\n'
                           '68. Development\n'
                           '开发\n'
                           'kāifā\n'
                           'Roots: 开 = open/start/เปิด + 发 = develop/emit/พัฒนา → development/การพัฒนา\n'
                           '\n'
                           '69. Error code\n'
                           '错误码\n'
                           'cuòwù mǎ\n'
                           'Roots: 错误 = error/ข้อผิดพลาด + 码 = code/รหัส → error code/รหัสข้อผิดพลาด\n'
                           '\n'
                           '70. Root Cause\n'
                           '根本原因\n'
                           'gēnběn yuányīn\n'
                           'Roots: 根本 = root/fundamental/รากฐาน + 原因 = cause/สาเหตุ → root '
                           'cause/สาเหตุราก\n'
                           '\n'
                           '71. Human-in-the-loop\n'
                           '人工审核\n'
                           'réngōng shěnhé\n'
                           'Roots: 人工 = human/manual/โดยคน + 审核 = review/check/ตรวจทาน → human '
                           'review/HITL\n'
                           '\n'
                           '72. Developer\n'
                           '开发人员\n'
                           'kāifā rényuán\n'
                           'Roots: 开发 = development/พัฒนา + 人员 = personnel/บุคลากร → '
                           'developer/นักพัฒนา\n'
                           '\n'
                           '73. Call\n'
                           '调用接口\n'
                           'diàoyòng jiēkǒu\n'
                           'Roots: 调用 = call/invoke / เรียกใช้งาน + 接口 = interface/API / API/จุดเชื่อม '
                           '→ Call / เรียก API\n'
                           '\n'
                           '74. Data Validation\n'
                           '数据验证\n'
                           'shùjù yànzhèng\n'
                           'Roots: 数据 = data / ข้อมูล + 验证 = validation / การตรวจยืนยัน → Data '
                           'Validation / ตรวจสอบข้อมูล\n'
                           '\n'
                           '75. Test Case\n'
                           '测试用例\n'
                           'cèshì yònglì\n'
                           'Roots: 测试 = testing / การทดสอบ + 用例 = use case / กรณีใช้งาน → Test Case / '
                           'กรณีทดสอบ\n'
                           '\n'
                           '76. UAT\n'
                           '验收测试\n'
                           'yànshōu cèshì\n'
                           'Roots: 验收 = acceptance / การตรวจรับ + 测试 = testing / การทดสอบ → UAT / '
                           'การทดสอบการยอมรับ'},
               {'chapter_id': 'chinese-tpm-ai-llm-platform',
                'chapter_title': 'AI / LLM / Platform',
                'order': 4,
                'content': 'AI / LLM / Platform\n'
                           '\n'
                           'How to use this chapter\n'
                           'Read the English first and try to recall the Chinese before looking at the '
                           'next line.\n'
                           'Then check the Pinyin and use the word-root note as a memory hook.\n'
                           'The root notes are learning aids based on modern word components; they are '
                           'not claiming to be historical etymology in every case.\n'
                           '\n'
                           '1. Currently\n'
                           '目前\n'
                           'mùqián\n'
                           'Roots: 目 = eye/current point/จุดที่มอง + 前 = before/front/ตรงหน้า → '
                           'currently/ปัจจุบัน\n'
                           '\n'
                           '2. Delay\n'
                           '延迟\n'
                           'yánchí\n'
                           'Roots: 延 = extend/delay/ยืด + 迟 = late/ช้า → latency/delay/ความหน่วง\n'
                           '\n'
                           '3. support\n'
                           '支持\n'
                           'zhīchí\n'
                           'Roots: 支 = support/ค้ำ + 持 = hold/ถือ → support/รองรับ\n'
                           '\n'
                           '4. every\n'
                           '每\n'
                           'měi\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 每 = every / ทุก; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '5. temperature\n'
                           '温度\n'
                           'wēndù\n'
                           'Roots: 温 = warm/อุณหภูมิ + 度 = degree/ระดับ → temperature/อุณหภูมิ\n'
                           '\n'
                           '6. set, configure\n'
                           '设置\n'
                           'shèzhì\n'
                           'Roots: 设 = set/ตั้ง + 置 = place/configure/วาง → set/configure/ตั้งค่า\n'
                           '\n'
                           '7. Context\n'
                           '上下文\n'
                           'shàngxiàwén\n'
                           'Roots: 上下 = above-below/before-after/รอบข้าง + 文 = text/ข้อความ → '
                           'context/บริบท\n'
                           '\n'
                           '8. too\n'
                           '太\n'
                           'tài\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 太 = too / เกินไป; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '9. use\n'
                           '用\n'
                           'yòng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 用 = use / ใช้; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '10. Format\n'
                           '格式\n'
                           'géshì\n'
                           'Roots: 格 = pattern/frame/กรอบ + 式 = form/style/รูปแบบ → format/รูปแบบ\n'
                           '\n'
                           '11. prevent\n'
                           '防止\n'
                           'fángzhǐ\n'
                           'Roots: 防 = prevent/ป้องกัน + 止 = stop/หยุด → prevent/ป้องกัน\n'
                           '\n'
                           '12. Prompt\n'
                           '提示词\n'
                           'tíshìcí\n'
                           'Roots: 提示 = prompt/hint/ชี้นำ + 词 = word/คำ → prompt/คำสั่งชี้นำ\n'
                           '\n'
                           '13. Injection\n'
                           '注入\n'
                           'zhùrù\n'
                           'Roots: 注 = pour/inject/ฉีดใส่ + 入 = enter/เข้า → injection/แทรกเข้า\n'
                           '\n'
                           '14. Embedding\n'
                           '嵌入\n'
                           'qiànrù\n'
                           'Roots: 嵌 = embed/inlay/ฝัง + 入 = enter/เข้า → embedding/ฝังเข้าไป\n'
                           '\n'
                           '15. split, chunk\n'
                           '切分\n'
                           'qiēfēn\n'
                           'Roots: 切 = cut/ตัด + 分 = divide/แบ่ง → split/chunk/แบ่งส่วน\n'
                           '\n'
                           '16. document\n'
                           '文档\n'
                           'wéndàng\n'
                           'Roots: 文 = text/document/ข้อความ + 档 = file/archive/แฟ้ม → document/เอกสาร\n'
                           '\n'
                           '17. Metadata\n'
                           '元数据\n'
                           'yuánshùjù\n'
                           'Roots: 元 = meta/origin/เหนือข้อมูล + 数据 = data/ข้อมูล → '
                           'metadata/ข้อมูลกำกับ\n'
                           '\n'
                           '18. Filter\n'
                           '过滤\n'
                           'guòlǜ\n'
                           'Roots: 过 = pass through/ผ่าน + 滤 = filter/กรอง → filter/กรองสิ่งที่ผ่าน\n'
                           '\n'
                           '19. condition\n'
                           '条件\n'
                           'tiáojiàn\n'
                           'Roots: 条 = item/condition/ข้อ + 件 = item/matter/รายการ → '
                           'condition/เงื่อนไข\n'
                           '\n'
                           '20. Retrieval\n'
                           '检索\n'
                           'jiǎnsuǒ\n'
                           'Roots: 检 = inspect/search/ตรวจค้น + 索 = search/ค้นหา → retrieval/การค้นคืน\n'
                           '\n'
                           '21. relevant, related\n'
                           '相关\n'
                           'xiāngguān\n'
                           'Roots: 相 = mutual/เกี่ยวกัน + 关 = relate/เชื่อม → related/เกี่ยวข้อง\n'
                           '\n'
                           '22. hybrid, mixed\n'
                           '混合\n'
                           'hùnhé\n'
                           'Roots: 混 = mix/ปน + 合 = combine/รวม → hybrid/mixed/ผสม\n'
                           '\n'
                           '23. Document\n'
                           '文件\n'
                           'wénjiàn\n'
                           'Roots: 文 = text/document/ข้อความ + 件 = item/piece/ชิ้น → document/เอกสาร\n'
                           '\n'
                           '24. Knowledge Base\n'
                           '知识库\n'
                           'zhīshikù\n'
                           'Roots: 知识 = knowledge/ความรู้ + 库 = repository/คลัง → knowledge '
                           'base/ฐานความรู้\n'
                           '\n'
                           '25. and\n'
                           '和\n'
                           'hé\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 和 = and / และ; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '26. answer\n'
                           '答案\n'
                           'dá’àn\n'
                           'Roots: 答 = answer/ตอบ + 案 = case/solution/คำตอบที่จัดไว้ → answer/คำตอบ\n'
                           '\n'
                           '27. Citation\n'
                           '引用\n'
                           'yǐnyòng\n'
                           'Roots: 引 = lead/draw/cite/ดึงอ้าง + 用 = use/ใช้ → '
                           'citation/reference/การอ้างอิง\n'
                           '\n'
                           '28. source\n'
                           '来源\n'
                           'láiyuán\n'
                           'Roots: 来 = come/มา + 源 = source/แหล่ง → source/origin/แหล่งที่มา\n'
                           '\n'
                           '29. Index\n'
                           '索引\n'
                           'suǒyǐn\n'
                           'Roots: 索 = search/ค้น + 引 = guide/นำทาง → index/ดัชนีช่วยค้น\n'
                           '\n'
                           '30. last, final\n'
                           '最后\n'
                           'zuìhòu\n'
                           'Roots: 最 = most / ที่สุด + 后 = after / หลัง → last, final / ล่าสุด, '
                           'สุดท้าย\n'
                           '\n'
                           '31. update\n'
                           '更新\n'
                           'gēngxīn\n'
                           'Roots: 更 = more / ยิ่ง, มากกว่า + 新 = new / ใหม่ → update / อัปเดต\n'
                           '\n'
                           '32. Fine-tuning\n'
                           '微调\n'
                           'wēitiáo\n'
                           'Roots: 微 = slight/fine/เล็กละเอียด + 调 = adjust/ปรับ → '
                           'fine-tuning/ปรับละเอียด\n'
                           '\n'
                           '33. toward, to\n'
                           '对\n'
                           'duì\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 对 = toward, to / ต่อ, กับ; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '34. benchmark, baseline\n'
                           '基准\n'
                           'jīzhǔn\n'
                           'Roots: 基 = base/foundation/ฐาน + 准 = standard/เกณฑ์ → '
                           'baseline/benchmark/เกณฑ์อ้างอิง\n'
                           '\n'
                           '35. Cost\n'
                           '成本\n'
                           'chéngběn\n'
                           'Roots: 成 = become/เกิดเป็น + 本 = base/cost/ต้นทุน → cost/ต้นทุน\n'
                           '\n'
                           '36. more\n'
                           '更\n'
                           'gèng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 更 = more / ยิ่ง, มากกว่า; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '37. high\n'
                           '高\n'
                           'gāo\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 高 = high / สูง; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '38. often\n'
                           '经常\n'
                           'jīngcháng\n'
                           'Roots: 经 = through/regular/เป็นประจำ + 常 = usual/ปกติ → often/บ่อย ๆ\n'
                           '\n'
                           '39. change\n'
                           '变化\n'
                           'biànhuà\n'
                           'Roots: 变 = change/เปลี่ยน + 化 = transform/กลาย → change/การเปลี่ยนแปลง\n'
                           '\n'
                           '40. therefore\n'
                           '所以\n'
                           'suǒyǐ\n'
                           'Roots: 所 = that which/สิ่งนั้น + 以 = therefore/by means of/จึง → '
                           'therefore/ดังนั้น\n'
                           '\n'
                           '41. suitable\n'
                           '适合\n'
                           'shìhé\n'
                           'Roots: 适 = suitable/เหมาะ + 合 = fit/match/เข้ากัน → suitable/เหมาะสม\n'
                           '\n'
                           '42. Evaluation\n'
                           '评估\n'
                           'pínggū\n'
                           'Roots: 评 = evaluate/ประเมิน + 估 = estimate/กะค่า → evaluation/การประเมิน\n'
                           '\n'
                           '43. Vendor / Supplier\n'
                           '供应商\n'
                           'gōngyìngshāng\n'
                           'Roots: 供应 = supply/จัดหา + 商 = merchant/vendor/ผู้ค้า → vendor/supplier\n'
                           '\n'
                           '44. Lock-in\n'
                           '锁定\n'
                           'suǒdìng\n'
                           'Roots: 锁 = lock/ล็อก + 定 = fix/set/ตรึง → lock-in/ถูกล็อกไว้\n'
                           '\n'
                           '45. answer\n'
                           '回答\n'
                           'huídá\n'
                           'Roots: 回 = return/ตอบกลับ + 答 = answer/ตอบ → answer/ตอบ\n'
                           '\n'
                           '46. accurate\n'
                           '准确\n'
                           'zhǔnquè\n'
                           'Roots: 准 = accurate/แม่น + 确 = certain/correct/ถูกต้อง → accurate/แม่นยำ\n'
                           '\n'
                           '47. Quality\n'
                           '质量\n'
                           'zhìliàng\n'
                           'Roots: 质 = quality/substance/คุณภาพ + 量 = measure/ปริมาณ → quality/คุณภาพ\n'
                           '\n'
                           '48. Evidence\n'
                           '依据\n'
                           'yījù\n'
                           'Roots: 依 = rely on/อาศัย + 据 = evidence/basis/หลักฐาน → evidence/basis\n'
                           '\n'
                           '49. Hallucination\n'
                           '幻觉\n'
                           'huànjué\n'
                           'Roots: 幻 = illusion/ภาพลวง + 觉 = perception/การรับรู้ → '
                           'hallucination/รับรู้สิ่งที่ไม่มีจริง\n'
                           '\n'
                           '50. rate\n'
                           '率\n'
                           'lǜ\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 率 = rate / อัตรา; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '51. below\n'
                           '低于\n'
                           'dīyú\n'
                           'Roots: 低 = low/ต่ำ + 于 = than/at/กว่า → below/lower than/ต่ำกว่า\n'
                           '\n'
                           '52. percent\n'
                           '百分之\n'
                           'bǎifēnzhī\n'
                           'Roots: 百 = hundred/ร้อย + 分 = part/ส่วน + 之 = of/ของ → percent/...ส่วนร้อย\n'
                           '\n'
                           '53. two\n'
                           '二\n'
                           'èr\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 二 = two / สอง; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '54. insufficient\n'
                           '不足\n'
                           'bùzú\n'
                           'Roots: 不 = not/ไม่ + 足 = enough/พอ → insufficient/ไม่เพียงพอ\n'
                           '\n'
                           '55. Refuse\n'
                           '拒绝\n'
                           'jùjué\n'
                           'Roots: 拒 = refuse/ปฏิเสธ + 绝 = cut off/ตัดขาด → refuse/ปฏิเสธ\n'
                           '\n'
                           '56. can\n'
                           '能\n'
                           'néng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 能 = can / สามารถ; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '57. bring\n'
                           '带来\n'
                           'dàilái\n'
                           'Roots: 带 = bring/carry/นำ + 来 = come/มา → bring/ก่อให้เกิด\n'
                           '\n'
                           '58. Value\n'
                           '价值\n'
                           'jiàzhí\n'
                           'Roots: 价 = price/value/ราคา + 值 = worth/value/คุณค่า → value/คุณค่า\n'
                           '\n'
                           '59. review, audit\n'
                           '审核\n'
                           'shěnhé\n'
                           'Roots: 审 = examine/ตรวจ + 核 = verify/core/ตรวจหลัก → review/audit/ตรวจทาน\n'
                           '\n'
                           '60. Assist\n'
                           '辅助\n'
                           'fǔzhù\n'
                           'Roots: 辅 = assist/support/ช่วยเสริม + 助 = help/ช่วย → assist/ช่วยเหลือ\n'
                           '\n'
                           '61. mode\n'
                           '模式\n'
                           'móshì\n'
                           'Roots: 模 = model/pattern/แบบ + 式 = form/style/รูปแบบ → mode/pattern/โหมด\n'
                           '\n'
                           '62. Start\n'
                           '开始\n'
                           'kāishǐ\n'
                           'Roots: 开 = open/start/เริ่ม + 始 = begin/เริ่มต้น → start/เริ่ม\n'
                           '\n'
                           '63. adopt\n'
                           '采用\n'
                           'cǎiyòng\n'
                           'Roots: 采 = select/adopt/เลือกใช้ + 用 = use/ใช้ → adopt/นำมาใช้\n'
                           '\n'
                           '64. Guardrail\n'
                           '护栏\n'
                           'hùlán\n'
                           'Roots: 护 = protect/ป้องกัน + 栏 = rail/barrier/ราวกั้น → '
                           'guardrail/ข้อกำกับป้องกัน\n'
                           '\n'
                           '65. end-to-\n'
                           '端到\n'
                           'duāndào\n'
                           'Roots: 端 = end/endpoint / ปลาย/ด้าน + 到 = to / ถึง, ไปยัง → end-to- / '
                           'จากปลายไปถึง...\n'
                           '\n'
                           '66. end\n'
                           '端\n'
                           'duān\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 端 = end / ปลาย, ด้าน; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '67. Tool\n'
                           '工具\n'
                           'gōngjù\n'
                           'Roots: 工 = work/งาน + 具 = tool/equipment/เครื่องมือ → tool/เครื่องมือ\n'
                           '\n'
                           '68. pass, through\n'
                           '通过\n'
                           'tōngguò\n'
                           'Roots: 通 = pass/connect/ผ่าน + 过 = pass/ผ่าน → pass/through/ผ่าน\n'
                           '\n'
                           '69. Sensitive\n'
                           '敏感\n'
                           'mǐngǎn\n'
                           'Roots: 敏 = quick/sensitive/ไว + 感 = feel/sense/รู้สึก → sensitive/อ่อนไหว\n'
                           '\n'
                           '70. send\n'
                           '发送\n'
                           'fāsòng\n'
                           'Roots: 发 = send/ปล่อย + 送 = send/deliver/ส่ง → send/ส่ง\n'
                           '\n'
                           '71. external\n'
                           '外部\n'
                           'wàibù\n'
                           'Roots: 外 = outside/ภายนอก + 部 = part/ส่วน → external/ภายนอก\n'
                           '\n'
                           '72. peak period\n'
                           '高峰期\n'
                           'gāofēngqī\n'
                           'Roots: 高峰 = peak/จุดสูงสุด + 期 = period/ช่วง → peak period/ช่วงพีค\n'
                           '\n'
                           '73. Concurrent\n'
                           '并发\n'
                           'bìngfā\n'
                           'Roots: 并 = together/พร้อมกัน + 发 = occur/send/เกิดหรือส่ง → '
                           'concurrent/ทำพร้อมกัน\n'
                           '\n'
                           '74. reduce\n'
                           '降低\n'
                           'jiàngdī\n'
                           'Roots: 降 = lower/ลด + 低 = low/ต่ำ → reduce/lower/ลดลง\n'
                           '\n'
                           '75. usage\n'
                           '使用量\n'
                           'shǐyòngliàng\n'
                           'Roots: 使用 = use/การใช้ + 量 = amount/ปริมาณ → usage/ปริมาณการใช้\n'
                           '\n'
                           '76. Cache\n'
                           '缓存\n'
                           'huǎncún\n'
                           'Roots: 缓 = slow/buffer/พัก + 存 = store/เก็บ → cache/เก็บพักไว้\n'
                           '\n'
                           '77. freshness\n'
                           '新鲜度\n'
                           'xīnxiāndù\n'
                           'Roots: 新鲜 = fresh/สดใหม่ + 度 = degree/ระดับ → freshness/ความสดใหม่\n'
                           '\n'
                           '78. occurrence, time\n'
                           '次\n'
                           'cì\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 次 = occurrence, time / ครั้ง; จำเป็นหน่วยคำตรง '
                           'ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '79. Experiment\n'
                           '实验\n'
                           'shíyàn\n'
                           'Roots: 实 = real/practical/จริง + 验 = test/ทดลองตรวจ → experiment/การทดลอง\n'
                           '\n'
                           '80. new\n'
                           '新\n'
                           'xīn\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 新 = new / ใหม่; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '81. shadow\n'
                           '影子\n'
                           'yǐngzi\n'
                           'Roots: 影 = shadow/เงา + 子 = suffix/object/คำลงท้าย → shadow/เงา\n'
                           '\n'
                           '82. Platform\n'
                           '平台\n'
                           'píngtái\n'
                           'Roots: 平 = flat/common/ราบกลาง + 台 = platform/stage/แท่น → '
                           'platform/ฐานกลาง\n'
                           '\n'
                           '83. personnel\n'
                           '人员\n'
                           'rényuán\n'
                           'Roots: 人 = person/คน + 员 = member/staff/บุคลากร → personnel/บุคลากร\n'
                           '\n'
                           '84. improve, increase\n'
                           '提高\n'
                           'tígāo\n'
                           'Roots: 提 = raise/ยก + 高 = high/สูง → improve/increase/ยกระดับ\n'
                           '\n'
                           '85. self-service\n'
                           '自助\n'
                           'zìzhù\n'
                           'Roots: 自 = self/ตัวเอง + 助 = help/ช่วย → self-help/ช่วยตัวเอง\n'
                           '\n'
                           '86. capability\n'
                           '能力\n'
                           'nénglì\n'
                           'Roots: 能 = ability/สามารถ + 力 = power/พลัง → capability/ความสามารถ\n'
                           '\n'
                           '87. only then\n'
                           '才\n'
                           'cái\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 才 = only then / จึงจะ, กว่าจะ; จำเป็นหน่วยคำตรง '
                           'ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '88. Template\n'
                           '模板\n'
                           'móbǎn\n'
                           'Roots: 模 = model/template/แบบ + 板 = board/plate/แม่แบบ → template/แม่แบบ\n'
                           '\n'
                           '89. team\n'
                           '团队\n'
                           'tuánduì\n'
                           'Roots: 团 = group/กลุ่ม + 队 = team/ทีม → team/ทีม\n'
                           '\n'
                           '90. independent\n'
                           '独立\n'
                           'dúlì\n'
                           'Roots: 独 = alone/เดี่ยว + 立 = stand/ยืน → independent/อิสระ\n'
                           '\n'
                           '91. Quota\n'
                           '配额\n'
                           'pèié\n'
                           'Roots: 配 = allocate/จัดสรร + 额 = quota/จำนวนกำหนด → quota/โควตา\n'
                           '\n'
                           '92. North Star\n'
                           '北极星\n'
                           'běijíxīng\n'
                           'Roots: 北 = north/เหนือ + 极 = pole/extreme/ขั้ว + 星 = star/ดาว → North '
                           'Star/ดาวเหนือ\n'
                           '\n'
                           '93. Pilot\n'
                           '试点\n'
                           'shìdiǎn\n'
                           'Roots: 试 = try/test/ทดลอง + 点 = point/site/จุด → pilot/ทดลองในจุดเล็กก่อน\n'
                           '\n'
                           '94. will\n'
                           '将\n'
                           'jiāng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 将 = will / จะ; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '95. one hundred\n'
                           '一百\n'
                           'yìbǎi\n'
                           'Roots: 一 = one/หนึ่ง + 百 = hundred/ร้อย → one hundred/หนึ่งร้อย\n'
                           '\n'
                           '96. measure word for people\n'
                           '名\n'
                           'míng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 名 = measure word for people / ลักษณนามคน; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '97. Significant\n'
                           '显著\n'
                           'xiǎnzhù\n'
                           'Roots: 显 = obvious/ชัด + 著 = marked/notable/เด่น → significant/มีนัยสำคัญ\n'
                           '\n'
                           '98. negative\n'
                           '负面\n'
                           'fùmiàn\n'
                           'Roots: 负 = negative/bear/ลบ + 面 = side/aspect/ด้าน → negative/เชิงลบ\n'
                           '\n'
                           '99. Feedback\n'
                           '反馈\n'
                           'fǎnkuì\n'
                           'Roots: 反 = back/reverse/ย้อน + 馈 = feed/give back/ป้อนกลับ → '
                           'feedback/ข้อมูลย้อนกลับ\n'
                           '\n'
                           '100. level, grade\n'
                           '等级\n'
                           'děngjí\n'
                           'Roots: 等 = class/equal/ชั้น + 级 = level/ระดับ → level/grade/ระดับ'},
               {'chapter_id': 'chinese-tpm-testing-uat-deployment',
                'chapter_title': 'Testing / UAT / Deployment',
                'order': 5,
                'content': 'Testing / UAT / Deployment\n'
                           '\n'
                           'How to use this chapter\n'
                           'Read the English first and try to recall the Chinese before looking at the '
                           'next line.\n'
                           'Then check the Pinyin and use the word-root note as a memory hook.\n'
                           'The root notes are learning aids based on modern word components; they are '
                           'not claiming to be historical etymology in every case.\n'
                           '\n'
                           '1. temporarily\n'
                           '暂时\n'
                           'zànshí\n'
                           'Roots: 暂 = temporary/ชั่วคราว + 时 = time/เวลา → temporarily/ชั่วคราว\n'
                           '\n'
                           '2. close, disable\n'
                           '关闭\n'
                           'guānbì\n'
                           'Roots: 关 = close/gate/ปิด + 闭 = shut/ปิด → close/disable/ปิด\n'
                           '\n'
                           '3. Function\n'
                           '功能\n'
                           'gōngnéng\n'
                           'Roots: 功 = function/achievement/หน้าที่ + 能 = ability/ความสามารถ → '
                           'function/ฟังก์ชัน\n'
                           '\n'
                           '4. related\n'
                           '有关\n'
                           'yǒuguān\n'
                           'Roots: 有 = have / มี + 关 = gate/pass / ด่าน/ประตู → related / เกี่ยวข้อง\n'
                           '\n'
                           '5. Responsible person\n'
                           '负责人\n'
                           'fùzérén\n'
                           'Roots: 负责 = be responsible/รับผิดชอบ + 人 = person/คน → responsible '
                           'person/ผู้รับผิดชอบ\n'
                           '\n'
                           '6. corrective rectification\n'
                           '整改\n'
                           'zhěnggǎi\n'
                           'Roots: 整 = put in order/จัดให้ถูก + 改 = change/correct/แก้ → '
                           'rectify/แก้ไขปรับปรุง\n'
                           '\n'
                           '7. measure, action\n'
                           '措施\n'
                           'cuòshī\n'
                           'Roots: 措 = arrange/take action/จัดการ + 施 = apply/ดำเนิน → '
                           'measure/action/มาตรการ\n'
                           '\n'
                           '8. today\n'
                           '今天\n'
                           'jīntiān\n'
                           'Roots: 今 = today/current/วันนี้ + 天 = day/วัน → today/วันนี้\n'
                           '\n'
                           '9. Decision\n'
                           '决定\n'
                           'juédìng\n'
                           'Roots: 决 = decide/ตัดสิน + 定 = fix/set/กำหนด → decision/การตัดสินใจ\n'
                           '\n'
                           '10. three\n'
                           '三\n'
                           'sān\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 三 = three / สาม; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '11. measure word\n'
                           '件\n'
                           'jiàn\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 件 = measure word / ลักษณนามสิ่ง/เรื่อง; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '12. matter, thing\n'
                           '事\n'
                           'shì\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 事 = matter, thing / เรื่อง, เหตุการณ์; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '13. explain\n'
                           '说明\n'
                           'shuōmíng\n'
                           'Roots: 说 = explain/say/พูด + 明 = clear/ชัด → explain/ชี้แจง\n'
                           '\n'
                           '14. current\n'
                           '当前\n'
                           'dāngqián\n'
                           'Roots: 当 = current/ตรงนี้ + 前 = before/front/หน้า → current/ปัจจุบัน\n'
                           '\n'
                           '15. behavior\n'
                           '行为\n'
                           'xíngwéi\n'
                           'Roots: 行 = act/do/กระทำ + 为 = do/be/ทำ → behavior/พฤติกรรม\n'
                           '\n'
                           '16. Expected\n'
                           '预期\n'
                           'yùqī\n'
                           'Roots: 预 = beforehand/ล่วงหน้า + 期 = expect/period/คาด → '
                           'expected/ที่คาดหวัง\n'
                           '\n'
                           '17. Reproduce\n'
                           '重现\n'
                           'chóngxiàn\n'
                           'Roots: 重 = again/อีกครั้ง + 现 = appear/ปรากฏ → reproduce/ทำให้เกิดซ้ำ\n'
                           '\n'
                           '18. replace, alternative\n'
                           '替代\n'
                           'tìdài\n'
                           'Roots: 替 = replace/แทน + 代 = substitute/generation/แทนที่ → '
                           'replace/alternative/ทดแทน\n'
                           '\n'
                           '19. ahead of time\n'
                           '提前\n'
                           'tíqián\n'
                           'Roots: 提 = move forward/ยกมา + 前 = ahead/ก่อน → ahead of time/ล่วงหน้า\n'
                           '\n'
                           '20. how long\n'
                           '多久\n'
                           'duōjiǔ\n'
                           'Roots: 多 = how much/มากเท่าไร + 久 = long time/นาน → how long/นานเท่าไร\n'
                           '\n'
                           '21. Delivery\n'
                           '交付\n'
                           'jiāofù\n'
                           'Roots: 交 = hand over/ส่งต่อ + 付 = deliver/pay/มอบ → delivery/การส่งมอบ\n'
                           '\n'
                           '22. be responsible\n'
                           '负责\n'
                           'fùzé\n'
                           'Roots: 负 = bear/รับภาระ + 责 = responsibility/ความรับผิดชอบ → be '
                           'responsible/รับผิดชอบ\n'
                           '\n'
                           '23. action\n'
                           '行动\n'
                           'xíngdòng\n'
                           'Roots: 行 = act/เดิน + 动 = move/เคลื่อน → action/การกระทำ\n'
                           '\n'
                           '24. item\n'
                           '项\n'
                           'xiàng\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 项 = item / รายการ, item; จำเป็นหน่วยคำตรง ๆ '
                           'ไม่ต้องฝืนแยก\n'
                           '\n'
                           '25. date\n'
                           '日期\n'
                           'rìqī\n'
                           'Roots: 日 = day/วัน + 期 = date/period/กำหนด → date/วันที่\n'
                           '\n'
                           '26. tomorrow\n'
                           '明天\n'
                           'míngtiān\n'
                           'Roots: 明 = next/bright/ถัดไป + 天 = day/วัน → tomorrow/พรุ่งนี้\n'
                           '\n'
                           '27. again\n'
                           '再\n'
                           'zài\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 再 = again / อีกครั้ง, แล้วค่อย; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '28. Follow up\n'
                           '跟进\n'
                           'gēnjìn\n'
                           'Roots: 跟 = follow/ตาม + 进 = advance/เดินหน้า → follow up/ติดตามต่อ\n'
                           '\n'
                           '29. choose\n'
                           '选择\n'
                           'xuǎnzé\n'
                           'Roots: 选 = choose/เลือก + 择 = select/คัด → choose/เลือก\n'
                           '\n'
                           '30. because\n'
                           '因为\n'
                           'yīnwèi\n'
                           'Roots: 因 = cause/เหตุ + 为 = because/as/เพราะ → because/เพราะว่า\n'
                           '\n'
                           '31. core\n'
                           '核心\n'
                           'héxīn\n'
                           'Roots: 核 = core/nucleus/แก่น + 心 = heart/ใจ → core/แกนหลัก\n'
                           '\n'
                           '32. before, front\n'
                           '前\n'
                           'qián\n'
                           'Roots: คำเดี่ยว/คำไวยากรณ์: 前 = before, front / ก่อน, ด้านหน้า; '
                           'จำเป็นหน่วยคำตรง ๆ ไม่ต้องฝืนแยก\n'
                           '\n'
                           '33. Security\n'
                           '安全\n'
                           'ānquán\n'
                           'Roots: 安 = safe/สงบ + 全 = complete/ครบ → security/safety/ความปลอดภัย\n'
                           '\n'
                           '34. review\n'
                           '审查\n'
                           'shěnchá\n'
                           'Roots: 审 = examine/ตรวจ + 查 = check/ตรวจ → review/ตรวจสอบ\n'
                           '\n'
                           '35. Interface\n'
                           '接口\n'
                           'jiēkǒu\n'
                           'Roots: 接 = connect/receive/เชื่อมรับ + 口 = opening/mouth/ช่อง → '
                           'interface/API/ช่องเชื่อม\n'
                           '\n'
                           '36. Query\n'
                           '查询\n'
                           'cháxún\n'
                           'Roots: 查 = check/search/ค้นตรวจ + 询 = inquire/สอบถาม → query/สืบค้น\n'
                           '\n'
                           '37. Consistency\n'
                           '一致性\n'
                           'yízhìxìng\n'
                           'Roots: 一致 = consistent/same/ตรงกัน + 性 = property/คุณสมบัติ → '
                           'consistency/ความสอดคล้อง\n'
                           '\n'
                           '38. NFR\n'
                           '非功能性需求\n'
                           'fēi gōngnéngxìng xūqiú\n'
                           'Roots: 非 = non-/ไม่ใช่ + 功能性 = functional/เชิงฟังก์ชัน + 需求 = requirement → '
                           'NFR\n'
                           '\n'
                           '39. Test Data\n'
                           '测试数据\n'
                           'cèshì shùjù\n'
                           'Roots: 测试 = testing/ทดสอบ + 数据 = data/ข้อมูล → test data/ข้อมูลทดสอบ\n'
                           '\n'
                           '40. Technical debt\n'
                           '技术债\n'
                           'jìshù zhài\n'
                           'Roots: 技术 = technical/เทคนิค + 债 = debt/หนี้ → technical '
                           'debt/หนี้ทางเทคนิค\n'
                           '\n'
                           '41. Sensitive data\n'
                           '敏感数据\n'
                           'mǐngǎn shùjù\n'
                           'Roots: 敏感 = sensitive/อ่อนไหว + 数据 = data/ข้อมูล → sensitive data\n'
                           '\n'
                           '42. Artificial\n'
                           '人工智能\n'
                           'réngōng zhìnéng\n'
                           'Roots: 人工 = artificial/human-made/มนุษย์สร้าง + 智能 = intelligence/ความฉลาด '
                           '→ AI\n'
                           '\n'
                           '43. Machine Learning\n'
                           '机器学习\n'
                           'jīqì xuéxí\n'
                           'Roots: 机器 = machine/เครื่อง + 学习 = learning/เรียนรู้ → machine learning\n'
                           '\n'
                           '44. Uncertainty\n'
                           '不确定性\n'
                           'bù quèdìngxìng\n'
                           'Roots: 不 = not/ไม่ + 确定 = certain/แน่นอน + 性 = state/property/ภาวะ → '
                           'uncertainty\n'
                           '\n'
                           '45. Throughput\n'
                           '吞吐量\n'
                           'tūntǔliàng\n'
                           'Roots: 吞 = swallow/in/รับเข้า + 吐 = spit/out/ส่งออก + 量 = amount/ปริมาณ → '
                           'throughput\n'
                           '\n'
                           '46. Temperature\n'
                           '温度参数\n'
                           'wēndù cānshù\n'
                           'Roots: 温度 = temperature/อุณหภูมิ + 参数 = parameter/พารามิเตอร์ → temperature '
                           'parameter\n'
                           '\n'
                           '47. Vector\n'
                           '向量\n'
                           'xiàngliàng\n'
                           'Roots: 向 = direction/ทิศทาง + 量 = quantity/ขนาด → vector/เวกเตอร์\n'
                           '\n'
                           '48. RAG\n'
                           '检索增强生成\n'
                           'jiǎnsuǒ zēngqiáng shēngchéng\n'
                           'Roots: 检索 = retrieval/ค้นคืน + 增强 = augment/เสริม + 生成 = generate/สร้าง → '
                           'RAG\n'
                           '\n'
                           '49. Benchmark\n'
                           '基准测试\n'
                           'jīzhǔn cèshì\n'
                           'Roots: 基准 = benchmark/baseline/เกณฑ์อ้างอิง + 测试 = test/ทดสอบ → benchmark '
                           'test\n'
                           '\n'
                           '50. Accuracy\n'
                           '准确率\n'
                           'zhǔnquèlǜ\n'
                           'Roots: 准确 = accurate/แม่นยำ + 率 = rate/อัตรา → accuracy rate\n'
                           '\n'
                           '51. Adoption\n'
                           '采用率\n'
                           'cǎiyòng lǜ\n'
                           'Roots: 采用 = adopt/use/นำมาใช้ + 率 = rate/อัตรา → adoption rate\n'
                           '\n'
                           '52. End-to-end\n'
                           '端到端\n'
                           'duāndào duān\n'
                           'Roots: 端 = end/ปลาย + 到 = to/ถึง + 端 = end/อีกปลาย → end-to-end/ต้นจนจบ\n'
                           '\n'
                           '53. Peak\n'
                           '高峰\n'
                           'gāofēng\n'
                           'Roots: 高 = high/สูง + 峰 = peak/summit/ยอด → peak/จุดสูงสุด\n'
                           '\n'
                           '54. Model Version\n'
                           '模型版本\n'
                           'móxíng bǎnběn\n'
                           'Roots: 模型 = model/โมเดล + 版本 = version/เวอร์ชัน → model version\n'
                           '\n'
                           '55. Shadow test\n'
                           '影子测试\n'
                           'yǐngzi cèshì\n'
                           'Roots: 影子 = shadow/เงา + 测试 = test/ทดสอบ → shadow test\n'
                           '\n'
                           '56. Lifecycle\n'
                           '生命周期\n'
                           'shēngmìng zhōuqī\n'
                           'Roots: 生命 = life/ชีวิต + 周期 = cycle/วงรอบ → lifecycle/วงจรชีวิต\n'
                           '\n'
                           '57. Self-service\n'
                           '自助服务\n'
                           'zìzhù fúwù\n'
                           'Roots: 自助 = self-help/ช่วยตัวเอง + 服务 = service/บริการ → self-service\n'
                           '\n'
                           '58. Corrective action\n'
                           '整改措施\n'
                           'zhěnggǎi cuòshī\n'
                           'Roots: 整改 = rectify/correct/แก้ปรับ + 措施 = measure/action/มาตรการ → '
                           'corrective action\n'
                           '\n'
                           '59. Alternative\n'
                           '替代方案\n'
                           'tìdài fāng’àn\n'
                           'Roots: 替代 = replace/alternative/ทดแทน + 方案 = plan/solution/แนวทาง → '
                           'alternative solution\n'
                           '\n'
                           '60. Core metric\n'
                           '核心指标\n'
                           'héxīn zhǐbiāo\n'
                           'Roots: 核心 = core/แกนหลัก + 指标 = metric/ตัวชี้วัด → core metric\n'
                           '\n'
                           '61. Security review\n'
                           '安全审查\n'
                           'ānquán shěnchá\n'
                           'Roots: 安全 = security/ความปลอดภัย + 审查 = review/ตรวจสอบ → security review\n'
                           '\n'
                           '62. Table\n'
                           '数据表\n'
                           'shùjùbiǎo\n'
                           'Roots: 数据 = data / ข้อมูล + 表 = table / ตาราง → Table / ตาราง\n'
                           '\n'
                           '63. Queue\n'
                           '队列\n'
                           'duìliè\n'
                           'Roots: 队 = team/queue unit / แถว/กลุ่ม + 列 = row/column/list / '
                           'แถว/คอลัมน์/ลำดับ → Queue / คิว\n'
                           '\n'
                           '64. Network\n'
                           '网络\n'
                           'wǎngluò\n'
                           'Roots: 网 = net/network / เครือข่าย + 络 = connect/network / เชื่อมโยง → '
                           'Network / เครือข่าย\n'
                           '\n'
                           '65. Alert\n'
                           '告警\n'
                           'gàojǐng\n'
                           'Roots: 告 = notify/tell / แจ้ง + 警 = warn/alert / เตือน → Alert / แจ้งเตือน\n'
                           '\n'
                           '66. Availability\n'
                           '可用性\n'
                           'kěyòngxìng\n'
                           'Roots: 可用 = available / พร้อมใช้งาน + 性 = property/-ness / คุณสมบัติ/ภาวะ → '
                           'Availability / ความพร้อมใช้งาน\n'
                           '\n'
                           '67. Scalability\n'
                           '可扩展性\n'
                           'kě kuòzhǎnxìng\n'
                           'Roots: 可 = can/สามารถ + 扩展 = scale/expand/ขยาย + 性 = '
                           'property/-ness/คุณสมบัติ → scalability/สเกลได้\n'
                           '\n'
                           '68. Business Requirement\n'
                           '业务需求\n'
                           'yèwù xūqiú\n'
                           'Roots: 业务 = business / ธุรกิจ + 需求 = requirement / ความต้องการ → Business '
                           'Requirement / ความต้องการทางธุรกิจ\n'
                           '\n'
                           '69. Requirement Change\n'
                           '需求变更\n'
                           'xūqiú biàngēng\n'
                           'Roots: 需求 = requirement / ความต้องการ + 变更 = Change / การเปลี่ยนแปลง → '
                           'Requirement Change / การเปลี่ยน Requirement\n'
                           '\n'
                           '70. Priority\n'
                           '优先级\n'
                           'yōuxiānjí\n'
                           'Roots: 优先 = priority/first / ก่อน/สำคัญก่อน + 级 = level / ระดับ → Priority '
                           '/ ลำดับความสำคัญ\n'
                           '\n'
                           '71. Business Process\n'
                           '业务流程\n'
                           'yèwù liúchéng\n'
                           'Roots: 业务 = business / ธุรกิจ + 流程 = process/flow / กระบวนการ → Business '
                           'Process / กระบวนการทางธุรกิจ\n'
                           '\n'
                           '72. Business Logic\n'
                           '业务逻辑\n'
                           'yèwù luójí\n'
                           'Roots: 业务 = business / ธุรกิจ + 逻辑 = logic / ตรรกะ → Business Logic / '
                           'ตรรกะทางธุรกิจ\n'
                           '\n'
                           '73. Use Case\n'
                           '使用场景\n'
                           'shǐyòng chǎngjǐng\n'
                           'Roots: 使用 = use / ใช้ + 场景 = Scenario / สถานการณ์ → Use Case / กรณีใช้งาน\n'
                           '\n'
                           '74. Acceptance Criteria\n'
                           '验收标准\n'
                           'yànshōu biāozhǔn\n'
                           'Roots: 验收 = acceptance / การตรวจรับ + 标准 = standard / มาตรฐาน → Acceptance '
                           'Criteria / เกณฑ์การยอมรับ\n'
                           '\n'
                           '75. Code\n'
                           '代码\n'
                           'dàimǎ\n'
                           'Roots: 代 = represent/substitute/แทน + 码 = code/รหัส → code/โค้ด\n'
                           '\n'
                           '76. Logic\n'
                           '逻辑\n'
                           'luójí\n'
                           'Roots: 逻辑 is a transliterated/loanword-style term for “logic”; '
                           'ควรจำเป็นก้อน = ตรรกะ มากกว่าฝืนแยกตัวอักษร\n'
                           '\n'
                           '77. Modify\n'
                           '修改\n'
                           'xiūgǎi\n'
                           'Roots: 修 = repair/modify / แก้/ซ่อม + 改 = change/modify / แก้ไข → Modify / '
                           'แก้ไข\n'
                           '\n'
                           '78. System Architecture\n'
                           '系统架构\n'
                           'xìtǒng jiàgòu\n'
                           'Roots: 系统 = System / ระบบ + 架构 = Architecture / สถาปัตยกรรม → System '
                           'Architecture / สถาปัตยกรรมระบบ\n'
                           '\n'
                           '79. Microservice\n'
                           '微服务\n'
                           'wēi fúwù\n'
                           'Roots: 微 = micro/small / เล็ก/ไมโคร + 服务 = service / บริการ → Microservice '
                           '/ ไมโครเซอร์วิส\n'
                           '\n'
                           '80. Component\n'
                           '组件\n'
                           'zǔjiàn\n'
                           'Roots: 组 = group/component / กลุ่ม + 件 = item/piece / ชิ้น/รายการ → '
                           'Component / ส่วนประกอบ\n'
                           '\n'
                           '81. Configuration\n'
                           '配置\n'
                           'pèizhì\n'
                           'Roots: 配 = configure/allocate / จัด/กำหนด + 置 = place/set / วาง/ตั้ง → '
                           'Configuration / การกำหนดค่า\n'
                           '\n'
                           '82. Pass Parameter\n'
                           '传参数\n'
                           'chuán cānshù\n'
                           'Roots: 传 = pass, transmit / ส่ง + 参数 = Parameter / พารามิเตอร์ → Pass '
                           'Parameter / ส่งพารามิเตอร์\n'
                           '\n'
                           '83. Integration\n'
                           '集成\n'
                           'jíchéng\n'
                           'Roots: 集 = gather/รวม + 成 = form/complete/เป็นผล → '
                           'integrate/integration/รวมให้เป็นระบบเดียว\n'
                           '\n'
                           '84. Data Transmission\n'
                           '传输数据\n'
                           'chuánshū shùjù\n'
                           'Roots: 传输 = transmit / ส่งผ่าน + 数据 = data / ข้อมูล → Data Transmission / '
                           'ส่งข้อมูล\n'
                           '\n'
                           '85. Timeout\n'
                           '超时\n'
                           'chāoshí\n'
                           'Roots: 超 = exceed/over / เกิน + 时 = time / เวลา → Timeout / หมดเวลา\n'
                           '\n'
                           '86. API Gateway\n'
                           '网关\n'
                           'wǎngguān\n'
                           'Roots: 网 = net/network / เครือข่าย + 关 = gate/pass / ด่าน/ประตู → API '
                           'Gateway / ประตู API\n'
                           '\n'
                           '87. Endpoint\n'
                           '端点\n'
                           'duāndiǎn\n'
                           'Roots: 端 = end/endpoint / ปลาย/ด้าน + 点 = point / จุด → Endpoint / ปลายทาง\n'
                           '\n'
                           '88. Authentication\n'
                           '身份验证\n'
                           'shēnfèn yànzhèng\n'
                           'Roots: 身份 = identity / ตัวตน + 验证 = validation / การตรวจยืนยัน → '
                           'Authentication / ยืนยันตัวตน\n'
                           '\n'
                           '89. Authorization\n'
                           '授权\n'
                           'shòuquán\n'
                           'Roots: 授 = grant / มอบ + 权 = right/permission / สิทธิ์ → Authorization / '
                           'การให้สิทธิ์\n'
                           '\n'
                           '90. Source Data\n'
                           '源数据\n'
                           'yuán shùjù\n'
                           'Roots: 源 = source / ต้นทาง/แหล่ง + 数据 = data / ข้อมูล → Source Data / '
                           'ข้อมูลต้นทาง\n'
                           '\n'
                           '91. Mapping\n'
                           '映射\n'
                           'yìngshè\n'
                           'Roots: 映 = reflect/map / ฉาย/สะท้อน + 射 = project/map / ฉาย/ยิง → Mapping / '
                           'การแมป\n'
                           '\n'
                           '92. Data Mismatch\n'
                           '数据不一致\n'
                           'shùjù bù yízhì\n'
                           'Roots: 数据 = data / ข้อมูล + 不 = not / ไม่ + 一致 = consistent / ตรงกัน → Data '
                           'Mismatch / ข้อมูลไม่ตรงกัน\n'
                           '\n'
                           '93. Data\n'
                           '数据一致性\n'
                           'shùjù yízhìxìng\n'
                           'Roots: 数据 = data / ข้อมูล + 一致性 = Consistency / ความสอดคล้อง → Data / '
                           'ความสอดคล้องของข้อมูล\n'
                           '\n'
                           '94. Data Migration\n'
                           '数据迁移\n'
                           'shùjù qiānyí\n'
                           'Roots: 数据 = data / ข้อมูล + 迁移 = Migration / การย้ายข้อมูล/ระบบ → Data '
                           'Migration / การย้ายข้อมูล\n'
                           '\n'
                           '95. Integration Test\n'
                           '集成测试\n'
                           'jíchéng cèshì\n'
                           'Roots: 集成 = integration / การผสานรวม + 测试 = testing / การทดสอบ → '
                           'Integration Test / การทดสอบการเชื่อมระบบ\n'
                           '\n'
                           '96. Regression Test\n'
                           '回归测试\n'
                           'huíguī cèshì\n'
                           'Roots: 回归 = regression / รีเกรสชัน + 测试 = testing / การทดสอบ → Regression '
                           'Test / การทดสอบถดถอย\n'
                           '\n'
                           '97. Performance Test\n'
                           '性能测试\n'
                           'xìngnéng cèshì\n'
                           'Roots: 性能 = performance / ประสิทธิภาพ + 测试 = testing / การทดสอบ → '
                           'Performance Test / การทดสอบประสิทธิภาพ\n'
                           '\n'
                           '98. Pass Testing\n'
                           '通过测试\n'
                           'tōngguò cèshì\n'
                           'Roots: 通过 = pass / ผ่าน + 测试 = testing / การทดสอบ → Pass Testing / '
                           'ผ่านการทดสอบ\n'
                           '\n'
                           '99. Fail Testing\n'
                           '测试未通过\n'
                           'cèshì wèi tōngguò\n'
                           'Roots: 测试 = testing / การทดสอบ + 未 = not yet/not / ยังไม่/ไม่ + 通过 = pass / '
                           'ผ่าน → Fail Testing / ไม่ผ่านการทดสอบ\n'
                           '\n'
                           '100. Development Environment\n'
                           '开发环境\n'
                           'kāifā huánjìng\n'
                           'Roots: 开发 = development / การพัฒนา + 环境 = environment / สภาพแวดล้อม → '
                           'Development Environment / สภาพแวดล้อมพัฒนา\n'
                           '\n'
                           '101. Test Environment\n'
                           '测试环境\n'
                           'cèshì huánjìng\n'
                           'Roots: 测试 = testing / การทดสอบ + 环境 = environment / สภาพแวดล้อม → Test '
                           'Environment / สภาพแวดล้อมทดสอบ\n'
                           '\n'
                           '102. Production Environment\n'
                           '生产环境\n'
                           'shēngchǎn huánjìng\n'
                           'Roots: 生产 = production / โปรดักชัน/การผลิต + 环境 = environment / สภาพแวดล้อม '
                           '→ Production Environment / สภาพแวดล้อมจริง\n'
                           '\n'
                           '103. Take Offline\n'
                           '下线\n'
                           'xiàxiàn\n'
                           'Roots: 下 = down/off/ลง + 线 = line/system/เส้นระบบ → take '
                           'offline/เอาลงจากระบบ\n'
                           '\n'
                           '104. Deployment Plan\n'
                           '部署计划\n'
                           'bùshǔ jìhuà\n'
                           'Roots: 部署 = Deployment / การนำขึ้นระบบ + 计划 = plan / แผน → Deployment Plan '
                           '/ แผน Deployment\n'
                           '\n'
                           '105. Release\n'
                           '发布窗口\n'
                           'fābù chuāngkǒu\n'
                           'Roots: 发布 = release / ปล่อยเวอร์ชัน + 窗口 = window / ช่วงเวลา/หน้าต่าง → '
                           'Release / ช่วงเวลาปล่อยระบบ\n'
                           '\n'
                           '106. Error Log\n'
                           '错误日志\n'
                           'cuòwù rìzhì\n'
                           'Roots: 错误 = error / ข้อผิดพลาด + 日志 = log / บันทึก → Error Log / '
                           'บันทึกข้อผิดพลาด\n'
                           '\n'
                           '107. Recover\n'
                           '恢复\n'
                           'huīfù\n'
                           'Roots: 恢 = restore/recover/คืนกลับ + 复 = return/repeat/กลับ → '
                           'recover/กู้คืน\n'
                           '\n'
                           '108. Workaround\n'
                           '临时解决方案\n'
                           "línshí jiějué fāng'àn\n"
                           'Roots: 临时 = temporary / ชั่วคราว + 解决方案 = solution / แนวทางแก้ปัญหา → '
                           'Workaround / วิธีแก้ชั่วคราว\n'
                           '\n'
                           '109. Preventive Action\n'
                           '预防措施\n'
                           'yùfáng cuòshī\n'
                           'Roots: 预防 = preventive / ป้องกัน + 措施 = measure/action / มาตรการ → '
                           'Preventive Action / มาตรการป้องกัน\n'
                           '\n'
                           '110. Completed\n'
                           '已完成\n'
                           'yǐ wánchéng\n'
                           'Roots: 已 = already / แล้ว + 完成 = complete / เสร็จสมบูรณ์ → Completed / '
                           'เสร็จแล้ว\n'
                           '\n'
                           '111. In Progress\n'
                           '进行中\n'
                           'jìnxíng zhōng\n'
                           'Roots: 进行 = in progress/carry out / ดำเนินการ + 中 = in/middle / อยู่ระหว่าง '
                           '→ In Progress / กำลังดำเนินการ\n'
                           '\n'
                           '112. Not Started\n'
                           '尚未开始\n'
                           'shàngwèi kāishǐ\n'
                           'Roots: 尚未 = not yet / ยังไม่ + 开始 = start / เริ่ม → Not Started / '
                           'ยังไม่ได้เริ่ม\n'
                           '\n'
                           '113. Waiting\n'
                           '等待\n'
                           'děngdài\n'
                           'Roots: 等 = wait/equal/รอ + 待 = await/คอย → waiting/รอ\n'
                           '\n'
                           '114. Blocked\n'
                           '阻塞\n'
                           'zǔsè\n'
                           'Roots: 阻 = block/ขวาง + 塞 = plug/block/อุด → blocked/ติดขัด\n'
                           '\n'
                           '115. Postpone\n'
                           '延期\n'
                           'yánqī\n'
                           'Roots: 延 = extend/ยืด + 期 = period/date/กำหนดเวลา → postpone/delay '
                           'date/เลื่อนกำหนด\n'
                           '\n'
                           '116. Estimate\n'
                           '预计\n'
                           'yùjì\n'
                           'Roots: 预 = beforehand/ล่วงหน้า + 计 = calculate/estimate/คำนวณ → '
                           'estimate/คาดการณ์\n'
                           '\n'
                           '117. Progress\n'
                           '进度\n'
                           'jìndù\n'
                           'Roots: 进 = advance/เดินหน้า + 度 = degree/progress level/ระดับ → '
                           'progress/ความคืบหน้า\n'
                           '\n'
                           '118. Deadline\n'
                           '截止日期\n'
                           'jiézhǐ rìqī\n'
                           'Roots: 截止 = cutoff/deadline / กำหนดสิ้นสุด + 日期 = date / วันที่ → Deadline '
                           '/ กำหนดส่ง\n'
                           '\n'
                           '119. Access Permission\n'
                           '访问权限\n'
                           'fǎngwèn quánxiàn\n'
                           'Roots: 访问 = access / เข้าถึง + 权限 = permission / สิทธิ์ → Access Permission '
                           '/ สิทธิ์การเข้าถึง\n'
                           '\n'
                           '120. Personal Data\n'
                           '个人数据\n'
                           'gèrén shùjù\n'
                           'Roots: 个人 = personal / ส่วนบุคคล + 数据 = data / ข้อมูล → Personal Data / '
                           'ข้อมูลส่วนบุคคล\n'
                           '\n'
                           '121. Cloud\n'
                           '云服务\n'
                           'yún fúwù\n'
                           'Roots: 云 = cloud / คลาวด์ + 服务 = service / บริการ → Cloud / ระบบคลาวด์\n'
                           '\n'
                           '122. Server\n'
                           '服务器\n'
                           'fúwùqì\n'
                           'Roots: 服务 = service/บริการ + 器 = device/machine/อุปกรณ์ → '
                           'server/เซิร์ฟเวอร์\n'
                           '\n'
                           '123. CI/CD\n'
                           '持续集成和持续交付\n'
                           'chíxù jíchéng hé chíxù jiāofù\n'
                           'Roots: 持续 = continuous / ต่อเนื่อง + 集成 = integration / การผสานรวม + 和 = '
                           'and / และ + 持续 = continuous / ต่อเนื่อง + 交付 = delivery / การส่งมอบ → CI/CD '
                           '/ การผสานรวมต่อเนื่องและการส่งมอบต่อเนื่อง\n'
                           '\n'
                           '124. Git Repository\n'
                           '代码仓库\n'
                           'dàimǎ cāngkù\n'
                           'Roots: 代码 = code / โค้ด + 仓库 = repository / คลัง → Git Repository / '
                           'คลังโค้ด\n'
                           '\n'
                           '125. Scale\n'
                           '扩展\n'
                           'kuòzhǎn\n'
                           'Roots: 扩 = expand/ขยาย + 展 = unfold/develop/กางออก → scale/expand/ขยาย\n'
                           '\n'
                           '126. LLM\n'
                           '大语言模型\n'
                           'dà yǔyán móxíng\n'
                           'Roots: 大 = large / ใหญ่ + 语言 = language / ภาษา + 模型 = model / โมเดล → LLM / '
                           'Large Language Model\n'
                           '\n'
                           '127. Embedding\n'
                           '向量嵌入\n'
                           'xiàngliàng qiànrù\n'
                           'Roots: 向量 = vector / เวกเตอร์ + 嵌入 = Embedding / เอ็มเบดดิง → Embedding / '
                           'การฝังเวกเตอร์\n'
                           '\n'
                           '128. Vector Database\n'
                           '向量数据库\n'
                           'xiàngliàng shùjùkù\n'
                           'Roots: 向量 = vector / เวกเตอร์ + 数据库 = database / ฐานข้อมูล → Vector '
                           'Database / ฐานข้อมูลเวกเตอร์\n'
                           '\n'
                           '129. Model Monitoring\n'
                           '模型监控\n'
                           'móxíng jiānkòng\n'
                           'Roots: 模型 = model / โมเดล + 监控 = monitoring / การเฝ้าระวัง → Model '
                           'Monitoring / การเฝ้าระวังโมเดล\n'
                           '\n'
                           '130. Privacy\n'
                           '隐私\n'
                           'yǐnsī\n'
                           'Roots: 隐 = hidden/private/ซ่อน + 私 = private/personal/ส่วนตัว → '
                           'privacy/ความเป็นส่วนตัว\n'
                           '\n'
                           '131. Cost per Request\n'
                           '单次请求成本\n'
                           'dāncì qǐngqiú chéngběn\n'
                           'Roots: 单次 = per occurrence/single request / ต่อครั้ง + 请求 = request / คำขอ '
                           '+ 成本 = cost / ต้นทุน → Cost per Request / ต้นทุนต่อคำขอ'}]},
 {'book_id': 'thinking-and-speaking',
  'title': 'Thinking & Speaking',
  'subtitle': 'Think Clearly. Speak Beautifully.',
  'author': 'Personal Growth Library',
  'content_type': 'Book',
  'category': 'Thinking & Communication',
  'description': 'Short reading chapters to build clearer thinking, elegant communication, structured '
                 'answers, and a distinctive personal voice.',
  'cover_emoji': '🧠',
  'audience': 'Mommy',
  'chapters': [{'chapter_id': 'thinking-speaking-01',
                'chapter_title': 'How to Form an Opinion',
                'order': 1,
                'content': 'A strong opinion is not the same as a fast opinion.\n'
                           '\n'
                           'People who sound thoughtful often do something quietly before they speak: '
                           'they separate what they know, what they assume, what they value, and what '
                           'they are still unsure about. That small separation makes their language '
                           'more precise and their thinking more credible.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Before asking, “What do I think?”, ask four smaller questions.\n'
                           '\n'
                           'What is the actual question?\n'
                           'What evidence do I have?\n'
                           'What principle matters to me here?\n'
                           'What would make me change my mind?\n'
                           '\n'
                           'This prevents a common mistake: answering a different question from the one '
                           'that was actually asked.\n'
                           '\n'
                           'Imagine someone asks, “Should every product team use AI?” A weak answer '
                           'jumps directly to yes or no. A stronger answer first defines the decision. '
                           'Are we talking about AI for customer-facing decisions, internal '
                           'productivity, or automation? Are accuracy and explainability critical? Is '
                           'the current problem actually caused by a lack of intelligence, or by a '
                           'broken process?\n'
                           '\n'
                           'Now an opinion can emerge from reasoning rather than instinct.\n'
                           '\n'
                           'A useful structure is:\n'
                           '\n'
                           'My current view is…\n'
                           'The main reason is…\n'
                           'The strongest evidence or example is…\n'
                           'The limitation is…\n'
                           'So in this situation, I would…\n'
                           '\n'
                           'Notice the phrase “my current view.” It sounds confident without pretending '
                           'that your thinking can never change. Intellectual maturity is not stubborn '
                           'certainty. It is the ability to hold a position and still remain responsive '
                           'to better evidence.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Ordinary:\n'
                           '“I think this is a bad idea.”\n'
                           '\n'
                           'More thoughtful:\n'
                           '“My concern is not the idea itself; it is the assumption underneath it.”\n'
                           '\n'
                           'Ordinary:\n'
                           '“I disagree.”\n'
                           '\n'
                           'More thoughtful:\n'
                           '“I see the logic, but I’m weighing the trade-off differently.”\n'
                           '\n'
                           'Ordinary:\n'
                           '“This will not work.”\n'
                           '\n'
                           'More thoughtful:\n'
                           '“I’m not convinced the current design addresses the constraint that is most '
                           'likely to determine the outcome.”\n'
                           '\n'
                           'These versions are not longer merely to sound sophisticated. They reveal '
                           'the reasoning behind the conclusion.\n'
                           '\n'
                           'A useful habit is to distinguish observation from interpretation.\n'
                           '\n'
                           'Observation:\n'
                           '“Customer completion dropped from 72% to 58% after the new step.”\n'
                           '\n'
                           'Interpretation:\n'
                           '“The new step may be introducing friction.”\n'
                           '\n'
                           'Conclusion:\n'
                           '“I would investigate the new step before redesigning the entire journey.”\n'
                           '\n'
                           'When these three layers are mixed together, people argue about conclusions '
                           'without realizing that they may actually disagree about the evidence or the '
                           'interpretation.\n'
                           '\n'
                           'Another habit is to find the strongest counterargument to your own view. Do '
                           'not choose a weak opposing argument just so you can defeat it. Ask, “If an '
                           'intelligent person disagreed with me, what would they say?” Then answer '
                           'that version.\n'
                           '\n'
                           'This makes you less defensive and more persuasive because you show that you '
                           'have looked beyond your preferred answer.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Choose one belief you hold about leadership, technology, education, or '
                           'work. Write one sentence for each:\n'
                           '\n'
                           'What I believe.\n'
                           'Why I believe it.\n'
                           'What evidence supports it.\n'
                           'What evidence could weaken it.\n'
                           'What I still do not know.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Speak for 90 seconds on this question:\n'
                           '\n'
                           '“What makes someone a good leader?”\n'
                           '\n'
                           'Do not begin with a list of adjectives. Begin with a point of view.\n'
                           '\n'
                           'Try:\n'
                           '“To me, leadership becomes visible when…”\n'
                           '\n'
                           'Then give one reason, one example, and one limitation.\n'
                           '\n'
                           'The goal is not to sound impressive. The goal is to make your thinking '
                           'visible.'},
               {'chapter_id': 'thinking-speaking-02',
                'chapter_title': 'How to Explain a Complex Idea Simply',
                'order': 2,
                'content': 'Explaining something simply is not the same as making it shallow.\n'
                           '\n'
                           'The real test of understanding is whether you can preserve the important '
                           'idea while removing unnecessary complexity. If an explanation becomes '
                           'confusing the moment technical vocabulary is removed, the speaker may know '
                           'the terminology without fully owning the concept.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Start with the listener, not with the topic.\n'
                           '\n'
                           'Before explaining, ask:\n'
                           'What do they already know?\n'
                           'What decision do they need to make?\n'
                           'Which detail changes that decision?\n'
                           'Which detail can wait?\n'
                           '\n'
                           'Experts often begin where their own knowledge begins. Good communicators '
                           'begin where the listener is.\n'
                           '\n'
                           'Suppose you need to explain an API to a business stakeholder.\n'
                           '\n'
                           'Technical-first explanation:\n'
                           '“An API exposes endpoints that allow clients to send requests and receive '
                           'structured responses.”\n'
                           '\n'
                           'Useful, but not always accessible.\n'
                           '\n'
                           'Listener-first explanation:\n'
                           '“Think of the API as the agreed doorway between two systems. One system '
                           'asks for something in a defined format, and the other system sends back a '
                           'defined answer. If we change that agreement carelessly, the systems can '
                           'stop understanding each other.”\n'
                           '\n'
                           'The second explanation is not technically perfect in every detail, but it '
                           'creates a mental model. Once the listener has the model, you can add '
                           'endpoint, request, response, authentication, timeout, and error handling.\n'
                           '\n'
                           'Use the ladder:\n'
                           '\n'
                           'One sentence → mental model → example → necessary detail → implication.\n'
                           '\n'
                           'For example:\n'
                           '\n'
                           'One sentence:\n'
                           '“RAG lets an AI answer using documents we choose.”\n'
                           '\n'
                           'Mental model:\n'
                           '“It is like letting the model open the right reference book before it '
                           'answers.”\n'
                           '\n'
                           'Example:\n'
                           '“If policy changes every week, we update the documents rather than '
                           'retraining the whole model.”\n'
                           '\n'
                           'Necessary detail:\n'
                           '“The system retrieves relevant chunks and gives them to the model as '
                           'context.”\n'
                           '\n'
                           'Implication:\n'
                           '“So retrieval quality and document freshness become product risks, not just '
                           'model quality.”\n'
                           '\n'
                           'That progression keeps the audience oriented.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of:\n'
                           '“It’s complicated.”\n'
                           '\n'
                           'Try:\n'
                           '“There are three moving parts; the interaction between them is what creates '
                           'the complexity.”\n'
                           '\n'
                           'Instead of:\n'
                           '“Basically…”\n'
                           '\n'
                           'Try:\n'
                           '“The simplest way to think about it is…”\n'
                           '\n'
                           'Instead of:\n'
                           '“You need to understand the architecture.”\n'
                           '\n'
                           'Try:\n'
                           '“The important part for this decision is how the data moves through the '
                           'architecture.”\n'
                           '\n'
                           'Good simplification also uses contrast.\n'
                           '\n'
                           '“Authentication answers ‘Who are you?’ Authorization answers ‘What are you '
                           'allowed to do?’”\n'
                           '\n'
                           '“Deployment means putting a version into an environment. Release means '
                           'making the capability available to users. They can happen together, but '
                           'they do not have to.”\n'
                           '\n'
                           'Contrast helps the brain create boundaries.\n'
                           '\n'
                           'Avoid using analogies that are memorable but misleading. Every analogy has '
                           'a breaking point. State it when necessary.\n'
                           '\n'
                           'For example:\n'
                           '“An API is like a waiter carrying requests between a customer and a '
                           'kitchen. That analogy is useful for request and response, but real APIs '
                           'also have authentication, contracts, retries, and machine-to-machine '
                           'behavior that the waiter analogy does not capture.”\n'
                           '\n'
                           'That sentence shows both clarity and intellectual discipline.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Pick one concept you understand well: credit decisioning, core banking, '
                           'RAG, UAT, or migration.\n'
                           '\n'
                           'Explain it in:\n'
                           'one sentence,\n'
                           'three sentences,\n'
                           'and one minute.\n'
                           '\n'
                           'If the one-minute version contains ten new terms, simplify again.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Answer:\n'
                           '“What is a Technical Product Manager?”\n'
                           '\n'
                           'Assume the listener is smart but has never worked in technology.\n'
                           '\n'
                           'Your target is not to prove how much you know. Your target is to make the '
                           'listener understand.'},
               {'chapter_id': 'thinking-speaking-03',
                'chapter_title': 'How to Structure Your Thoughts Before Speaking',
                'order': 3,
                'content': 'Many people think good speakers think faster. Often they simply structure '
                           'earlier.\n'
                           '\n'
                           'When your mind contains ten ideas at once, speech becomes difficult because '
                           'every idea feels equally important. Structure creates hierarchy. It tells '
                           'you what comes first, what supports it, and what can be left out.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Use a four-part spine:\n'
                           '\n'
                           'Conclusion → Reason → Example → Reflection.\n'
                           '\n'
                           'Conclusion tells the listener where you are going.\n'
                           'Reason explains why.\n'
                           'Example makes the idea concrete.\n'
                           'Reflection shows judgment.\n'
                           '\n'
                           'Question:\n'
                           '“How do you handle a production incident?”\n'
                           '\n'
                           'Unstructured:\n'
                           '“I check logs, talk to IT, see customer impact, maybe do workaround, then '
                           'find root cause, then communicate…”\n'
                           '\n'
                           'Structured:\n'
                           '“My first priority is to contain customer impact before optimizing for a '
                           'perfect diagnosis. I separate the work into containment, diagnosis, and '
                           'permanent correction. For example, if applications stop receiving results, '
                           'I first confirm the affected population and establish a workaround, while '
                           'the technical team traces the failing layer. After recovery, I make sure '
                           'the permanent fix closes the root cause and that monitoring can detect the '
                           'same pattern earlier next time.”\n'
                           '\n'
                           'The content is similar. The hierarchy is different.\n'
                           '\n'
                           'Another useful structure is:\n'
                           '\n'
                           'Situation → Tension → Decision → Result → Learning.\n'
                           '\n'
                           'This works well for interviews because it prevents a career story from '
                           'becoming a chronology.\n'
                           '\n'
                           'Situation gives context.\n'
                           'Tension explains why the situation mattered.\n'
                           'Decision shows judgment.\n'
                           'Result shows consequence.\n'
                           'Learning shows growth.\n'
                           '\n'
                           'Do not confuse structure with sounding robotic. The structure should live '
                           'underneath the answer, not appear as labels in every conversation.\n'
                           '\n'
                           'You do not have to say, “My first reason is…” unless the context is formal. '
                           'You can sound natural:\n'
                           '\n'
                           '“What mattered most was…”\n'
                           '“The reason I approached it that way was…”\n'
                           '“One example was…”\n'
                           '“What I learned from that was…”\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of starting with background:\n'
                           '“So, first of all, at that time our team was…”\n'
                           '\n'
                           'Start with the point:\n'
                           '“The hardest part was not the system defect; it was making a decision with '
                           'incomplete information.”\n'
                           '\n'
                           'Then give the background needed to understand that point.\n'
                           '\n'
                           'Instead of:\n'
                           '“There are many things.”\n'
                           '\n'
                           'Try:\n'
                           '“I would separate this into three decisions.”\n'
                           '\n'
                           'Instead of:\n'
                           '“It depends.”\n'
                           '\n'
                           'Try:\n'
                           '“It depends mainly on two variables…”\n'
                           '\n'
                           '“It depends” is often correct, but unfinished. Name what it depends on.\n'
                           '\n'
                           'A useful speaking discipline is the headline test. Before answering, '
                           'silently complete:\n'
                           '\n'
                           '“If they remember only one sentence from my answer, I want it to be…”\n'
                           '\n'
                           'That sentence becomes your opening.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Take a question you often receive at work. Write ten points you could '
                           'mention.\n'
                           '\n'
                           'Now force yourself to choose only three.\n'
                           '\n'
                           'Then identify which one is the conclusion and which two are support.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Answer in two minutes:\n'
                           '\n'
                           '“Tell me about a difficult problem you solved.”\n'
                           '\n'
                           'Use Conclusion → Reason → Example → Reflection.\n'
                           '\n'
                           'Record yourself once. Listen only for structure, not accent or grammar.\n'
                           '\n'
                           'Ask:\n'
                           'Could someone summarize my point after hearing me once?'},
               {'chapter_id': 'thinking-speaking-04',
                'chapter_title': 'How to Disagree Elegantly',
                'order': 4,
                'content': 'Elegant disagreement protects two things at the same time: the quality of '
                           'the decision and the dignity of the people involved.\n'
                           '\n'
                           'Weak disagreement attacks the person.\n'
                           'Avoidant disagreement hides the problem.\n'
                           'Strong disagreement makes the difference in reasoning visible.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Separate the person from the proposition.\n'
                           '\n'
                           'Do not think:\n'
                           '“They are wrong.”\n'
                           '\n'
                           'Think:\n'
                           '“We are using different assumptions, evidence, priorities, or risk '
                           'tolerances.”\n'
                           '\n'
                           'That shift changes your language.\n'
                           '\n'
                           'A practical sequence is:\n'
                           '\n'
                           'Acknowledge → Locate the difference → Explain your reasoning → Invite '
                           'examination.\n'
                           '\n'
                           'Example:\n'
                           '\n'
                           '“I agree that speed matters here. Where I see it differently is the '
                           'rollback risk. If we release the full scope now, a data issue would be '
                           'difficult to reverse. I would rather reduce scope and preserve a clean '
                           'recovery path. Is there a constraint I’m missing that makes the full '
                           'release necessary?”\n'
                           '\n'
                           'This is not soft. It is precise.\n'
                           '\n'
                           'Notice that “I agree that speed matters” does not mean you agree with the '
                           'proposal. It shows that you understand the value the other person is '
                           'protecting.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of:\n'
                           '“No, that’s not right.”\n'
                           '\n'
                           'Try:\n'
                           '“I’m looking at the same information and reaching a different conclusion.”\n'
                           '\n'
                           'Instead of:\n'
                           '“You don’t understand the problem.”\n'
                           '\n'
                           'Try:\n'
                           '“I think we may be framing the problem at different levels.”\n'
                           '\n'
                           'Instead of:\n'
                           '“That makes no sense.”\n'
                           '\n'
                           'Try:\n'
                           '“I can follow the logic up to this point; the assumption I’m not yet '
                           'comfortable with is…”\n'
                           '\n'
                           'Instead of:\n'
                           '“We can’t do that.”\n'
                           '\n'
                           'Try:\n'
                           '“We can do it, but the cost is X. The decision is whether that trade-off is '
                           'acceptable.”\n'
                           '\n'
                           'Good disagreement also knows when to become firmer.\n'
                           '\n'
                           'If there is a safety, regulatory, ethical, or material customer risk, '
                           'excessive politeness can blur responsibility.\n'
                           '\n'
                           'You can say:\n'
                           '\n'
                           '“I want to be explicit: I do not recommend proceeding without this control '
                           'because the failure mode affects customer funds.”\n'
                           '\n'
                           'That is clear without being insulting.\n'
                           '\n'
                           'Another principle is to argue at the right level.\n'
                           '\n'
                           'If someone proposes Feature A and you prefer Feature B, you may waste time '
                           'comparing features when the real disagreement is about the objective.\n'
                           '\n'
                           'Ask:\n'
                           '“What outcome are we optimizing for?”\n'
                           '“What risk are we trying to reduce?”\n'
                           '“What assumption would have to be true for this option to work?”\n'
                           '\n'
                           'Those questions move the conversation from preference to reasoning.\n'
                           '\n'
                           'Do not use questions as disguised attacks. “Do you really think that makes '
                           'sense?” is not curiosity. It is criticism wearing a question mark.\n'
                           '\n'
                           'Real curiosity sounds different:\n'
                           '“Can you walk me through what makes that risk acceptable in this case?”\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Recall a recent disagreement.\n'
                           '\n'
                           'Was the difference mainly about:\n'
                           'facts,\n'
                           'assumptions,\n'
                           'values,\n'
                           'priority,\n'
                           'risk tolerance,\n'
                           'or timing?\n'
                           '\n'
                           'How would the conversation have changed if you had named the real '
                           'difference?\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Practice this sentence:\n'
                           '\n'
                           '“I see the benefit of that approach. My hesitation is…”\n'
                           '\n'
                           'Then explain one trade-off in under 60 seconds.\n'
                           '\n'
                           'Your goal is not to win the conversation. Your goal is to make the decision '
                           'better.'},
               {'chapter_id': 'thinking-speaking-05',
                'chapter_title': 'How to Sound Thoughtful, Not Rehearsed',
                'order': 5,
                'content': 'Prepared is good. Rehearsed can sound lifeless.\n'
                           '\n'
                           'The difference is whether you have memorized sentences or understood '
                           'ideas.\n'
                           '\n'
                           'When people memorize full scripts, one unexpected question can break the '
                           'sequence. When they understand the structure and the meaning, they can '
                           'rebuild the answer in real time.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Memorize anchors, not paragraphs.\n'
                           '\n'
                           'For an interview story, your anchors might be:\n'
                           '\n'
                           'Problem.\n'
                           'Why it mattered.\n'
                           'Decision.\n'
                           'Trade-off.\n'
                           'Result.\n'
                           'Learning.\n'
                           '\n'
                           'You do not need to memorize every connecting sentence. The anchors keep you '
                           'oriented while your language stays natural.\n'
                           '\n'
                           'Thoughtful speech also contains small signs of real-time reasoning:\n'
                           '\n'
                           '“The way I think about it is…”\n'
                           '“What mattered most in that situation was…”\n'
                           '“There are two parts to that.”\n'
                           '“My first instinct would be X, but I would check Y before deciding.”\n'
                           '“I used to think X; experience made me pay more attention to Y.”\n'
                           '\n'
                           'These phrases work because they expose the path of thought.\n'
                           '\n'
                           'But do not fill every answer with hedging.\n'
                           '\n'
                           'Too much:\n'
                           '“Maybe, perhaps, I guess, kind of, probably…”\n'
                           '\n'
                           'Balanced:\n'
                           '“My current view is X, mainly because of Y. The part I would want to '
                           'validate is Z.”\n'
                           '\n'
                           'Confidence and uncertainty can coexist.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Rehearsed:\n'
                           '“I am a highly motivated and results-oriented professional with extensive '
                           'experience…”\n'
                           '\n'
                           'Thoughtful:\n'
                           '“The pattern across my work is that I tend to step into problems that sit '
                           'between business and technology, especially when the path is unclear.”\n'
                           '\n'
                           'Rehearsed:\n'
                           '“My weakness is that I am a perfectionist.”\n'
                           '\n'
                           'Thoughtful:\n'
                           '“One behavior I’ve had to change is staying too close to execution after '
                           'the team already understands the direction. As my scope grows, I need to '
                           'create clarity and then give people room to own the work.”\n'
                           '\n'
                           'The second version sounds more credible because it contains a real '
                           'mechanism.\n'
                           '\n'
                           'Use specific verbs. “Managed” can mean almost anything.\n'
                           '\n'
                           'Instead of:\n'
                           '“I managed the issue.”\n'
                           '\n'
                           'Try:\n'
                           '“I framed the decision, aligned the owners, set the validation criteria, '
                           'and escalated the dependency that was blocking recovery.”\n'
                           '\n'
                           'Specificity creates authority without needing dramatic adjectives.\n'
                           '\n'
                           'Another technique is controlled imperfection. Spoken language does not need '
                           'to sound like an essay. Short pauses, a sentence restart, or “The key point '
                           'is…” can make you sound more present.\n'
                           '\n'
                           'Do not optimize every sentence while you are speaking. Optimize the '
                           'thought.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Choose one interview answer you have practiced many times.\n'
                           '\n'
                           'Reduce it to five anchor words.\n'
                           '\n'
                           'Now speak it three times without trying to repeat the same wording.\n'
                           '\n'
                           'If the meaning stays stable while the sentences change, you own the '
                           'answer.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Answer:\n'
                           '\n'
                           '“What is one thing your manager would say about you?”\n'
                           '\n'
                           'Do not use generic adjectives first.\n'
                           '\n'
                           'Start with a pattern:\n'
                           '“She would probably say that when a problem is ambiguous, I tend to…”\n'
                           '\n'
                           'Then give evidence.\n'
                           '\n'
                           'A good answer should feel familiar to you but new in the moment.'},
               {'chapter_id': 'thinking-speaking-06',
                'chapter_title': 'How to Tell a Story With a Point',
                'order': 6,
                'content': 'A story is not a list of events.\n'
                           '\n'
                           'A useful story changes the listener’s understanding of something: your '
                           'judgment, a problem, a relationship, or a lesson. If nothing changes, the '
                           'story may be accurate but it will feel flat.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Every professional story needs tension.\n'
                           '\n'
                           'Tension is the gap between:\n'
                           'what was happening,\n'
                           'and what needed to happen.\n'
                           '\n'
                           'For example:\n'
                           '\n'
                           '“The system was technically live, but customers were not receiving the '
                           'application result. The team could fix defects one by one, but each '
                           'recurrence cost us another day of volume.”\n'
                           '\n'
                           'Now the listener knows why the story matters.\n'
                           '\n'
                           'A strong story can follow:\n'
                           '\n'
                           'Context → Tension → Choice → Consequence → Meaning.\n'
                           '\n'
                           'Context should be short.\n'
                           'Tension should be clear.\n'
                           'Choice should show agency.\n'
                           'Consequence should show what changed.\n'
                           'Meaning should answer, “Why are you telling me this?”\n'
                           '\n'
                           'Compare:\n'
                           '\n'
                           'Flat:\n'
                           '“There was a system issue. We had meetings with IT. We checked logs. Then '
                           'it was fixed.”\n'
                           '\n'
                           'With a point:\n'
                           '“The recurring issue taught me that incident management cannot stop at '
                           'technical recovery. We were fixing each symptom in about two days, but the '
                           'business kept absorbing the same type of loss. I changed the discussion '
                           'from ‘When can we fix this case?’ to ‘What pattern are these cases '
                           'revealing, and what control would prevent the next one?’ That shifted the '
                           'work from case-by-case recovery toward a more durable operating model.”\n'
                           '\n'
                           'The second story contains meaning.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of:\n'
                           '“Then this happened, then this happened…”\n'
                           '\n'
                           'Use causal transitions:\n'
                           '“Because of that…”\n'
                           '“That created a second problem…”\n'
                           '“The turning point was…”\n'
                           '“So I had to choose between…”\n'
                           '“What changed after that was…”\n'
                           '\n'
                           'Causal language makes a story feel intelligent because events are '
                           'connected, not merely sequenced.\n'
                           '\n'
                           'Do not overcrowd stories with names, dates, system acronyms, and side '
                           'characters unless they matter to the point.\n'
                           '\n'
                           'Ask of every detail:\n'
                           'Does this help the listener understand the tension, the choice, or the '
                           'result?\n'
                           '\n'
                           'If not, remove it.\n'
                           '\n'
                           'The ending matters. Do not end with:\n'
                           '“And yeah, that was it.”\n'
                           '\n'
                           'End with the meaning:\n'
                           '“That experience changed how I run incidents now.”\n'
                           '“What I took from it is that clarity of ownership is itself a control.”\n'
                           '“It taught me that being collaborative does not mean avoiding hard '
                           'decisions.”\n'
                           '\n'
                           'A story can end in learning even if the result was not perfect.\n'
                           '\n'
                           'In fact, stories where everything went smoothly often reveal less judgment '
                           'than stories involving trade-offs or mistakes.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Choose one work experience.\n'
                           '\n'
                           'Write:\n'
                           'The situation in one sentence.\n'
                           'The tension in one sentence.\n'
                           'The decision you personally influenced.\n'
                           'The measurable or observable result.\n'
                           'The idea you learned.\n'
                           '\n'
                           'If the “decision” line says only “we discussed,” go deeper.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Tell the story in 90 seconds.\n'
                           '\n'
                           'Then tell it again in 45 seconds.\n'
                           '\n'
                           'The shorter version will reveal what the story is really about.'},
               {'chapter_id': 'thinking-speaking-07',
                'chapter_title': 'How to Answer Difficult Questions',
                'order': 7,
                'content': 'A difficult question becomes harder when you believe you must answer '
                           'immediately.\n'
                           '\n'
                           'Strong communicators create a small amount of thinking space without '
                           'sounding evasive. They do not rush to fill silence with the first idea '
                           'available.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Use three moves:\n'
                           '\n'
                           'Clarify → Frame → Answer.\n'
                           '\n'
                           'Clarify when the question contains an ambiguity that changes the answer.\n'
                           '\n'
                           '“Do you mean the product decision or the technical implementation?”\n'
                           '\n'
                           'Frame when the question is broad.\n'
                           '\n'
                           '“I’d separate that into customer impact and execution risk.”\n'
                           '\n'
                           'Then answer.\n'
                           '\n'
                           'This turns a large question into a manageable one.\n'
                           '\n'
                           'If you genuinely do not know, do not bluff.\n'
                           '\n'
                           'A useful answer is:\n'
                           '\n'
                           '“I don’t know that number with confidence. What I would check is X, because '
                           'it determines Y.”\n'
                           '\n'
                           'That shows judgment without inventing facts.\n'
                           '\n'
                           'For hypothetical questions, state assumptions.\n'
                           '\n'
                           '“If we assume the deadline is fixed and security controls cannot be '
                           'reduced, then I would trade scope before quality.”\n'
                           '\n'
                           'The assumption makes your reasoning auditable.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of:\n'
                           '“I don’t know.”\n'
                           '\n'
                           'Try:\n'
                           '“I haven’t worked with that directly, so I wouldn’t want to overstate my '
                           'experience. My understanding is…, and the first thing I would validate '
                           'is…”\n'
                           '\n'
                           'Instead of:\n'
                           '“It depends.”\n'
                           '\n'
                           'Try:\n'
                           '“It depends mainly on the reversibility of the decision and the impact of '
                           'being wrong.”\n'
                           '\n'
                           'Instead of:\n'
                           '“That never happened to me.”\n'
                           '\n'
                           'Try:\n'
                           '“I haven’t had that exact scenario, but I’ve handled a closely related '
                           'problem where…”\n'
                           '\n'
                           'Do not force every question into a story. Sometimes the best answer is a '
                           'framework.\n'
                           '\n'
                           'Question:\n'
                           '“How would you prioritize five urgent issues?”\n'
                           '\n'
                           'Answer:\n'
                           '“I would first separate urgency from severity. I’d look at customer harm, '
                           'regulatory exposure, transaction value, number of users affected, and '
                           'whether the failure is still expanding. Then I’d check which action is '
                           'reversible and which issue has a safe workaround. That gives me an order I '
                           'can defend rather than relying on who is shouting the loudest.”\n'
                           '\n'
                           'That answer demonstrates thinking directly.\n'
                           '\n'
                           'When a question feels adversarial, lower the emotional temperature by '
                           'becoming more concrete.\n'
                           '\n'
                           '“What specifically concerns you about the approach?”\n'
                           '“Which outcome would make you consider it successful?”\n'
                           '“Is the concern cost, timing, or control?”\n'
                           '\n'
                           'Specificity turns conflict into a solvable problem.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Write down three questions that make you nervous.\n'
                           '\n'
                           'For each, identify:\n'
                           'What is ambiguous?\n'
                           'What framework could organize the answer?\n'
                           'What fact would you refuse to guess?\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Practice answering:\n'
                           '“What is a technical area where you are not yet strong?”\n'
                           '\n'
                           'Use:\n'
                           'truth,\n'
                           'current understanding,\n'
                           'how you learn,\n'
                           'and evidence of progress.\n'
                           '\n'
                           'You do not need to know everything. You need to show that you know how to '
                           'think when you do not.'},
               {'chapter_id': 'thinking-speaking-08',
                'chapter_title': 'How to Ask Intelligent Questions',
                'order': 8,
                'content': 'A good question does more than collect information. It improves the quality '
                           'of thought in the room.\n'
                           '\n'
                           'People often try to sound intelligent by asking complicated questions. The '
                           'better goal is to ask the question that reveals what matters.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'High-value questions usually do one of five things:\n'
                           '\n'
                           'Clarify the objective.\n'
                           'Expose an assumption.\n'
                           'Reveal a constraint.\n'
                           'Connect cause and effect.\n'
                           'Force a decision.\n'
                           '\n'
                           'Objective:\n'
                           '“What outcome are we trying to change?”\n'
                           '\n'
                           'Assumption:\n'
                           '“What has to be true for this solution to work?”\n'
                           '\n'
                           'Constraint:\n'
                           '“Which part is actually fixed: date, scope, budget, or control '
                           'requirement?”\n'
                           '\n'
                           'Cause and effect:\n'
                           '“What evidence tells us this change caused the drop?”\n'
                           '\n'
                           'Decision:\n'
                           '“What decision needs to be made today, and what information is still '
                           'missing?”\n'
                           '\n'
                           'These questions move work forward.\n'
                           '\n'
                           'Before asking a technical question, know why you need the answer.\n'
                           '\n'
                           'Weak:\n'
                           '“What database do we use?”\n'
                           '\n'
                           'Stronger:\n'
                           '“Which system is the source of truth for this field? I’m trying to '
                           'understand which team should own the correction.”\n'
                           '\n'
                           'The second question connects technical information to a product decision.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of:\n'
                           '“Why did you do that?”\n'
                           '\n'
                           'Try:\n'
                           '“What trade-off led you to that design?”\n'
                           '\n'
                           'Instead of:\n'
                           '“Can this be faster?”\n'
                           '\n'
                           'Try:\n'
                           '“Which component dominates the current latency, and what would we give up '
                           'to reduce it?”\n'
                           '\n'
                           'Instead of:\n'
                           '“Is it safe?”\n'
                           '\n'
                           'Try:\n'
                           '“What failure mode are we most concerned about, and what control contains '
                           'it?”\n'
                           '\n'
                           'Good questions avoid hidden accusations.\n'
                           '\n'
                           '“Why didn’t you test this?” assumes negligence.\n'
                           '\n'
                           '“What coverage did we have before release, and which scenario was missing?” '
                           'investigates the system.\n'
                           '\n'
                           'Intelligent questions also recognize timing. During an incident, '
                           'philosophical questions can wait. Ask facts first:\n'
                           '\n'
                           'What is affected?\n'
                           'When did it start?\n'
                           'Is impact still growing?\n'
                           'What changed?\n'
                           'What is the safest containment?\n'
                           'Who owns the next decision?\n'
                           '\n'
                           'After stabilization, ask deeper questions about process and prevention.\n'
                           '\n'
                           'A senior question often connects levels.\n'
                           '\n'
                           '“How does this technical limitation change the customer experience?”\n'
                           '“If we accept this workaround for three months, what operational debt are '
                           'we creating?”\n'
                           '“What would scale differently if volume doubled?”\n'
                           '\n'
                           'Those questions show that you can move between detail and consequence.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'In your next meeting, do not aim to ask many questions.\n'
                           '\n'
                           'Aim to ask one question that changes the discussion.\n'
                           '\n'
                           'Before speaking, ask yourself:\n'
                           'Am I asking because I need to know, because the team needs to know, or '
                           'because I want to demonstrate knowledge?\n'
                           '\n'
                           'Only the first two help the meeting.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Take a simple request:\n'
                           '“We want real-time loan status.”\n'
                           '\n'
                           'Ask five questions that progressively deepen the problem.\n'
                           '\n'
                           'Start with the customer outcome.\n'
                           'End with the operational or technical risk.'},
               {'chapter_id': 'thinking-speaking-09',
                'chapter_title': 'How to Speak With Nuance',
                'order': 9,
                'content': 'Nuance is the ability to see that two apparently conflicting things can '
                           'both contain truth.\n'
                           '\n'
                           'It is not indecision. It is precision.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Avoid false binaries.\n'
                           '\n'
                           'Fast versus safe.\n'
                           'Customer value versus risk.\n'
                           'Strong leadership versus listening.\n'
                           'Automation versus human judgment.\n'
                           '\n'
                           'Real decisions often ask:\n'
                           'How much of each?\n'
                           'Under what conditions?\n'
                           'For which cases?\n'
                           'At what stage?\n'
                           '\n'
                           'Instead of:\n'
                           '“Automation is better.”\n'
                           '\n'
                           'Try:\n'
                           '“Automation creates the most value in high-volume, repeatable cases where '
                           'the decision rules are stable. For ambiguous or high-impact exceptions, '
                           'human review may still be the better control.”\n'
                           '\n'
                           'The nuanced version is more useful because it contains boundaries.\n'
                           '\n'
                           'A strong thinker also separates principle from context.\n'
                           '\n'
                           'Principle:\n'
                           '“Customers should receive a timely decision.”\n'
                           '\n'
                           'Context:\n'
                           '“A real-time decision may not be appropriate if a required risk check is '
                           'asynchronous.”\n'
                           '\n'
                           'Now you can protect the principle without pretending every implementation '
                           'should be identical.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Useful nuanced phrases:\n'
                           '\n'
                           '“Both can be true.”\n'
                           '“The distinction I would make is…”\n'
                           '“That is generally true, but the exception that matters here is…”\n'
                           '“I agree with the principle; I’m less certain about this implementation.”\n'
                           '“The benefit increases with X, but so does the risk.”\n'
                           '“In the short term…, while in the longer term…”\n'
                           '“For most cases…, but for high-risk cases…”\n'
                           '\n'
                           'These phrases are powerful because they add dimensions.\n'
                           '\n'
                           'Do not use nuance to avoid commitment. After exploring both sides, decide.\n'
                           '\n'
                           'For example:\n'
                           '“There is a genuine trade-off between speed and traceability. In this case, '
                           'because the decision affects credit eligibility, I would prioritize '
                           'traceability and accept the additional processing time.”\n'
                           '\n'
                           'That is nuanced and decisive.\n'
                           '\n'
                           'Another skill is proportion.\n'
                           '\n'
                           'Not every problem deserves the same emotional or operational response. '
                           'Ask:\n'
                           'How large is the impact?\n'
                           'How reversible is the error?\n'
                           'How confident are we?\n'
                           'How quickly can we learn?\n'
                           '\n'
                           'Nuanced people calibrate.\n'
                           '\n'
                           'They do not treat a UI typo like a financial miscalculation, and they do '
                           'not treat a regulatory control like a cosmetic preference.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Choose a belief you often state strongly.\n'
                           '\n'
                           'Write:\n'
                           'When is it true?\n'
                           'When is it not true?\n'
                           'What variable changes the answer?\n'
                           'What decision would you still make?\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Answer:\n'
                           '“Should leaders protect their teams?”\n'
                           '\n'
                           'Give a nuanced answer.\n'
                           '\n'
                           'Consider protection from unnecessary noise, accountability for real '
                           'mistakes, psychological safety, and the risk of shielding people from '
                           'useful feedback.\n'
                           '\n'
                           'End with your principle.'},
               {'chapter_id': 'thinking-speaking-10',
                'chapter_title': 'How to Use Examples and Analogies',
                'order': 10,
                'content': 'Abstract ideas are difficult because the listener has nothing to hold.\n'
                           '\n'
                           'Examples create evidence. Analogies create shape.\n'
                           '\n'
                           'Used well, both make your thinking easier to remember.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Use an example when you need to prove or demonstrate.\n'
                           'Use an analogy when you need to build a mental model.\n'
                           '\n'
                           'Example:\n'
                           '“A rollback plan matters. In a migration, if a schema change cannot be '
                           'reversed safely, the team may need a forward fix rather than a traditional '
                           'rollback.”\n'
                           '\n'
                           'Analogy:\n'
                           '“A rollback plan is like knowing the safe exit before you enter a tunnel.”\n'
                           '\n'
                           'The analogy is memorable. The example is precise. Together they are '
                           'stronger.\n'
                           '\n'
                           'A useful analogy maps one relationship, not the entire system.\n'
                           '\n'
                           'For RAG:\n'
                           '“Think of the model as a smart reader and retrieval as the librarian '
                           'choosing which pages to put on the desk.”\n'
                           '\n'
                           'This helps explain why a brilliant model can still answer badly if '
                           'retrieval brings the wrong pages.\n'
                           '\n'
                           'Then state the limit:\n'
                           '“Of course, a real RAG system also includes permissions, ranking, chunking, '
                           'freshness, and evaluation.”\n'
                           '\n'
                           'That prevents the analogy from becoming misinformation.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of:\n'
                           '“Let me give you an example.”\n'
                           '\n'
                           'Sometimes try:\n'
                           '“A concrete case makes this clearer.”\n'
                           '“You can see the trade-off in a simple example.”\n'
                           '“One situation where this matters is…”\n'
                           '\n'
                           'Instead of adding examples randomly, use them to answer a question:\n'
                           'What does this look like?\n'
                           'Why does it matter?\n'
                           'When does it fail?\n'
                           '\n'
                           'A good example contains only the detail needed for the point.\n'
                           '\n'
                           'If your point is about ownership, do not spend 40 seconds explaining system '
                           'names.\n'
                           '\n'
                           'If your point is about customer impact, include the impact.\n'
                           '\n'
                           'Examples also make interview answers credible. Generic claims become '
                           'evidence.\n'
                           '\n'
                           'Claim:\n'
                           '“I can lead without formal authority.”\n'
                           '\n'
                           'Evidence:\n'
                           '“I have led cross-functional sessions where no one reported to me. My role '
                           'was to frame the problem, make the dependencies visible, and get each owner '
                           'to commit to an action and date.”\n'
                           '\n'
                           'Now the listener can judge the claim.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Take three concepts:\n'
                           'API,\n'
                           'technical debt,\n'
                           'and leadership without authority.\n'
                           '\n'
                           'Create one analogy for each.\n'
                           '\n'
                           'Then write where each analogy breaks.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Explain “technical debt” to a non-technical executive in one minute.\n'
                           '\n'
                           'Use:\n'
                           'one sentence,\n'
                           'one analogy,\n'
                           'one business example,\n'
                           'and one implication.'},
               {'chapter_id': 'thinking-speaking-11',
                'chapter_title': 'How to Make Your Language More Elegant',
                'order': 11,
                'content': 'Elegant language is usually simpler than people expect.\n'
                           '\n'
                           'It is not created by replacing ordinary words with rare vocabulary. It '
                           'comes from precision, rhythm, contrast, and restraint.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Choose the strongest simple word.\n'
                           '\n'
                           'Instead of:\n'
                           '“utilize” → “use”\n'
                           '“in order to” → “to”\n'
                           '“due to the fact that” → “because”\n'
                           '“at this point in time” → “now”\n'
                           '\n'
                           'Clean language gives important ideas more space.\n'
                           '\n'
                           'Then improve the sentence structure.\n'
                           '\n'
                           'Flat:\n'
                           '“I learned a lot from the incident and it was difficult and many teams were '
                           'involved.”\n'
                           '\n'
                           'Better:\n'
                           '“The incident was difficult not because one system failed, but because no '
                           'single team could see the whole flow.”\n'
                           '\n'
                           'The second sentence has contrast and focus.\n'
                           '\n'
                           'Three techniques create elegance:\n'
                           '\n'
                           '1. Contrast\n'
                           '“Speed matters, but reversibility matters more when the cost of being wrong '
                           'is high.”\n'
                           '\n'
                           '2. Parallel structure\n'
                           '“We needed to contain the impact, understand the cause, and prevent '
                           'recurrence.”\n'
                           '\n'
                           '3. A precise closing line\n'
                           '“Clarity was not a communication benefit; it was part of the control.”\n'
                           '\n'
                           'These patterns make ideas memorable without sounding theatrical.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Ordinary:\n'
                           '“I like solving problems.”\n'
                           '\n'
                           'More distinctive:\n'
                           '“I’m most engaged when the problem is still ambiguous and the team needs a '
                           'way to turn uncertainty into a decision.”\n'
                           '\n'
                           'Ordinary:\n'
                           '“I work well with many teams.”\n'
                           '\n'
                           'More distinctive:\n'
                           '“A large part of my work is creating enough shared clarity that people who '
                           'do not report to me can still move in the same direction.”\n'
                           '\n'
                           'Ordinary:\n'
                           '“I’m responsible.”\n'
                           '\n'
                           'More distinctive:\n'
                           '“If I own the outcome, I stay with the problem until the decision, owner, '
                           'and next step are clear.”\n'
                           '\n'
                           'Notice that the upgraded versions describe behavior rather than '
                           'adjectives.\n'
                           '\n'
                           'Good language also uses rhythm. Mix shorter and longer sentences.\n'
                           '\n'
                           'Long sentence:\n'
                           '“I want the team to understand the objective, the constraint, and the '
                           'trade-off before we commit.”\n'
                           '\n'
                           'Short sentence:\n'
                           '“Then I let them own the how.”\n'
                           '\n'
                           'The short line lands because the previous sentence prepared it.\n'
                           '\n'
                           'Avoid decorating every sentence. If every line tries to be profound, none '
                           'of them feels natural.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Take five sentences you often use in interviews.\n'
                           '\n'
                           'Remove:\n'
                           'generic adjectives,\n'
                           'unnecessary intensifiers,\n'
                           'and vague verbs.\n'
                           '\n'
                           'Replace them with:\n'
                           'specific behavior,\n'
                           'a contrast,\n'
                           'or a concrete outcome.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Answer:\n'
                           '“What kind of leader do you want to be?”\n'
                           '\n'
                           'Use ordinary vocabulary.\n'
                           '\n'
                           'Aim for one sentence worth remembering.'},
               {'chapter_id': 'thinking-speaking-12',
                'chapter_title': 'How to Speak Like a Leader',
                'order': 12,
                'content': 'Leadership language shifts the focus from personal activity to collective '
                           'direction.\n'
                           '\n'
                           'This does not mean hiding your contribution. It means showing that your '
                           'contribution changes how other people can act.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Move from:\n'
                           '“I did the work.”\n'
                           '\n'
                           'Toward:\n'
                           '“I created clarity, made a decision, aligned owners, removed a constraint, '
                           'or enabled the team to execute.”\n'
                           '\n'
                           'Individual contributor language:\n'
                           '“I checked the issue and followed up with IT.”\n'
                           '\n'
                           'Leadership language:\n'
                           '“I framed the impact, aligned the technical and business owners on the '
                           'decision path, and escalated the dependency that was preventing recovery.”\n'
                           '\n'
                           'The second version still contains your work, but it explains the leverage.\n'
                           '\n'
                           'A leader speaks in outcomes and decision rights.\n'
                           '\n'
                           'What outcome matters?\n'
                           'Who owns which decision?\n'
                           'What does the team need from me?\n'
                           'What should I not take away from them?\n'
                           '\n'
                           'Strong leaders do not prove value by doing everyone’s job.\n'
                           '\n'
                           'They create context.\n'
                           '\n'
                           'For example:\n'
                           '“The team did not need me to tell them how to code the fix. They needed a '
                           'clear definition of acceptable customer impact, the priority order, and a '
                           'decision on whether we would accept a workaround.”\n'
                           '\n'
                           'That sentence shows role clarity.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of:\n'
                           '“I told them to…”\n'
                           '\n'
                           'Try:\n'
                           '“I set the direction that…”\n'
                           '“We aligned on…”\n'
                           '“I made the trade-off explicit…”\n'
                           '“I asked the owner to commit to…”\n'
                           '“I removed the blocker by…”\n'
                           '\n'
                           'Instead of:\n'
                           '“My team worked for me.”\n'
                           '\n'
                           'Try:\n'
                           '“I created enough alignment that each function could move without waiting '
                           'for repeated escalation.”\n'
                           '\n'
                           'Instead of:\n'
                           '“I protected my team from management.”\n'
                           '\n'
                           'Try:\n'
                           '“I absorb unnecessary noise, but I do not shield the team from information '
                           'they need to make good decisions.”\n'
                           '\n'
                           'That is a more mature version of protection.\n'
                           '\n'
                           'Leadership speech also gives credit accurately.\n'
                           '\n'
                           '“We solved it” can hide your role.\n'
                           '“I solved it” can erase the team.\n'
                           '\n'
                           'Try:\n'
                           '“My role was to frame the problem and drive the decision; the engineering '
                           'team designed and implemented the fix.”\n'
                           '\n'
                           'Now ownership is precise.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Choose one achievement from your career.\n'
                           '\n'
                           'Write two columns:\n'
                           'What I personally did.\n'
                           'What became possible for other people because I did it.\n'
                           '\n'
                           'The second column is where leadership often lives.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Answer:\n'
                           '“How would your role change if you had two strong team members reporting to '
                           'you?”\n'
                           '\n'
                           'Do not say only that you would assign tasks.\n'
                           '\n'
                           'Talk about context, quality of thinking, ownership, coaching, escalation, '
                           'and space.'},
               {'chapter_id': 'thinking-speaking-13',
                'chapter_title': 'How to Reflect on an Experience',
                'order': 13,
                'content': 'Experience does not automatically become wisdom.\n'
                           '\n'
                           'Reflection is the process that turns an event into a reusable idea.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'After something important happens, ask:\n'
                           '\n'
                           'What happened?\n'
                           'Why did it happen?\n'
                           'What did I believe at the time?\n'
                           'What turned out to be different?\n'
                           'What will I do differently next time?\n'
                           '\n'
                           'The most interesting line is often the fourth.\n'
                           '\n'
                           '“I believed the team needed more detailed instructions. What turned out to '
                           'be missing was a shared definition of the outcome.”\n'
                           '\n'
                           'That is growth.\n'
                           '\n'
                           'Reflection is different from self-criticism. The purpose is not to punish '
                           'your past self. It is to improve your model of the world.\n'
                           '\n'
                           'A useful reflection contains a before and after.\n'
                           '\n'
                           'Before:\n'
                           '“I thought leadership meant having the answer.”\n'
                           '\n'
                           'After:\n'
                           '“I now think leadership often means creating the conditions for the best '
                           'answer to emerge.”\n'
                           '\n'
                           'Before:\n'
                           '“I treated incidents as technical defects.”\n'
                           '\n'
                           'After:\n'
                           '“I now treat incidents as customer-impact events with technical causes.”\n'
                           '\n'
                           'The contrast makes development visible.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Instead of:\n'
                           '“I learned communication is important.”\n'
                           '\n'
                           'Try:\n'
                           '“I learned that communication during an incident is not a status-reporting '
                           'activity; it is part of risk control because different teams make decisions '
                           'from the information we provide.”\n'
                           '\n'
                           'Instead of:\n'
                           '“I should have done better.”\n'
                           '\n'
                           'Try:\n'
                           '“The decision was reasonable with the information we had, but the process '
                           'exposed one gap I would change next time…”\n'
                           '\n'
                           'Instead of:\n'
                           '“It was a failure.”\n'
                           '\n'
                           'Try:\n'
                           '“The outcome was below what we wanted. The useful question is which '
                           'assumption failed.”\n'
                           '\n'
                           'Reflection also makes success more useful.\n'
                           '\n'
                           'When something goes well, ask why.\n'
                           '\n'
                           'Was the result caused by your process, timing, luck, team capability, or a '
                           'favorable constraint?\n'
                           '\n'
                           'If you do not know, you cannot repeat it.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Choose:\n'
                           'one success,\n'
                           'one difficult experience,\n'
                           'and one mistake.\n'
                           '\n'
                           'For each, write:\n'
                           'what happened,\n'
                           'the hidden assumption,\n'
                           'what you know now,\n'
                           'and one behavior you changed.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Answer:\n'
                           '“Tell me about something you would handle differently today.”\n'
                           '\n'
                           'Do not spend most of the answer apologizing for the past.\n'
                           '\n'
                           'Spend it showing the upgrade in your thinking.'},
               {'chapter_id': 'thinking-speaking-14',
                'chapter_title': 'How to Develop Intellectual Curiosity',
                'order': 14,
                'content': 'Curiosity is not knowing many facts.\n'
                           '\n'
                           'It is the habit of noticing that an answer opens another question.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Use five lenses:\n'
                           '\n'
                           'Why?\n'
                           'What else?\n'
                           'What if?\n'
                           'Compared with what?\n'
                           'How would we know?\n'
                           '\n'
                           'Why?\n'
                           '“Why did conversion fall?”\n'
                           '\n'
                           'What else?\n'
                           '“What else changed at the same time?”\n'
                           '\n'
                           'What if?\n'
                           '“What if the problem is not the new screen but slower response time?”\n'
                           '\n'
                           'Compared with what?\n'
                           '“Is the drop large compared with normal weekly variation?”\n'
                           '\n'
                           'How would we know?\n'
                           '“What evidence would distinguish those explanations?”\n'
                           '\n'
                           'This sequence protects you from accepting the first plausible story.\n'
                           '\n'
                           'Curiosity also crosses disciplines.\n'
                           '\n'
                           'A product problem can be:\n'
                           'a psychology problem,\n'
                           'an incentive problem,\n'
                           'a data problem,\n'
                           'an operating-model problem,\n'
                           'or a technology problem.\n'
                           '\n'
                           'The more lenses you can use, the less likely you are to confuse the first '
                           'explanation with the whole explanation.\n'
                           '\n'
                           'Read outside your profession for this reason.\n'
                           '\n'
                           'History teaches consequence over time.\n'
                           'Psychology teaches behavior.\n'
                           'Economics teaches incentives and trade-offs.\n'
                           'Literature teaches motive and perspective.\n'
                           'Science teaches evidence and falsifiability.\n'
                           'Engineering teaches constraints and failure modes.\n'
                           '\n'
                           'You do not need to become an expert in all of them. You are building a '
                           'larger library of questions.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Curious:\n'
                           '“What am I missing?”\n'
                           '\n'
                           'More precise:\n'
                           '“What would have to be true for the opposite conclusion to be correct?”\n'
                           '\n'
                           'Curious:\n'
                           '“Why?”\n'
                           '\n'
                           'More precise:\n'
                           '“Which mechanism connects this change to that outcome?”\n'
                           '\n'
                           'Curious:\n'
                           '“Any other option?”\n'
                           '\n'
                           'More precise:\n'
                           '“If this constraint disappeared, what would we design differently?”\n'
                           '\n'
                           'Questions improve when they become testable.\n'
                           '\n'
                           'Curiosity also requires patience with not knowing.\n'
                           '\n'
                           'Do not rush to close the question because uncertainty is uncomfortable.\n'
                           '\n'
                           'Sometimes the intelligent sentence is:\n'
                           '“I have a hypothesis, but I don’t think we have enough evidence yet to call '
                           'it the root cause.”\n'
                           '\n'
                           'That protects quality.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Pick one problem at work.\n'
                           '\n'
                           'Write one question from each of the five lenses.\n'
                           '\n'
                           'Then choose the question whose answer would most change your decision.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Explain one opinion you changed in the last year.\n'
                           '\n'
                           'What new evidence, experience, or perspective changed it?\n'
                           '\n'
                           'Changing your mind for a good reason is evidence of thinking, not '
                           'weakness.'},
               {'chapter_id': 'thinking-speaking-15',
                'chapter_title': 'How to Develop Your Own Voice',
                'order': 15,
                'content': 'A strong voice is recognizable not because it imitates someone impressive, '
                           'but because it consistently reflects how you see the world.\n'
                           '\n'
                           'You can learn from speakers you admire: their clarity, calmness, structure, '
                           'vocabulary, or ability to connect ideas. But the goal is to borrow '
                           'principles, not a personality.\n'
                           '\n'
                           'CORE IDEA\n'
                           '\n'
                           'Your voice is built from three things:\n'
                           '\n'
                           'What you notice.\n'
                           'What you value.\n'
                           'How you connect ideas.\n'
                           '\n'
                           'Two people can describe the same project and sound completely different.\n'
                           '\n'
                           'One notices execution speed.\n'
                           'Another notices customer behavior.\n'
                           'Another notices power and incentives.\n'
                           'Another notices system design.\n'
                           '\n'
                           'Your voice becomes clearer when you know what you naturally pay attention '
                           'to.\n'
                           '\n'
                           'Look across your own stories.\n'
                           '\n'
                           'Do you repeatedly care about:\n'
                           'clarity,\n'
                           'fairness,\n'
                           'customer impact,\n'
                           'systems,\n'
                           'learning,\n'
                           'independence,\n'
                           'quality,\n'
                           'or bringing people together?\n'
                           '\n'
                           'Those recurring themes are not slogans. They are evidence of perspective.\n'
                           '\n'
                           'Then choose language that fits you.\n'
                           '\n'
                           'If you are naturally analytical, you do not need to become dramatic.\n'
                           'If you are naturally warm, you do not need to sound cold to sound senior.\n'
                           'If you think quickly, your development may be learning to slow the delivery '
                           'so other people can follow.\n'
                           '\n'
                           'The best voice feels like a more disciplined version of you.\n'
                           '\n'
                           'LANGUAGE UPGRADE\n'
                           '\n'
                           'Generic:\n'
                           '“I’m passionate about products.”\n'
                           '\n'
                           'Personal:\n'
                           '“I like products because they force me to connect three worlds at once: '
                           'what people need, what the business can sustain, and what the system can '
                           'reliably deliver.”\n'
                           '\n'
                           'Generic:\n'
                           '“I value teamwork.”\n'
                           '\n'
                           'Personal:\n'
                           '“I enjoy the moment when people who started with different assumptions can '
                           'finally see the same problem and move together.”\n'
                           '\n'
                           'Generic:\n'
                           '“I want to grow.”\n'
                           '\n'
                           'Personal:\n'
                           '“I want to keep expanding the range of problems I can understand deeply '
                           'enough to make good decisions.”\n'
                           '\n'
                           'Your voice strengthens through repetition with variation.\n'
                           '\n'
                           'Keep a small “idea notebook.”\n'
                           '\n'
                           'When you notice a thought worth keeping, write:\n'
                           'the idea,\n'
                           'why you believe it,\n'
                           'one example,\n'
                           'and one sentence you like.\n'
                           '\n'
                           'Over time, you will build your own language.\n'
                           '\n'
                           'THINK ABOUT IT\n'
                           '\n'
                           'Complete these sentences:\n'
                           '\n'
                           'I pay attention to…\n'
                           'I become frustrated when…\n'
                           'I respect people who…\n'
                           'I believe good work should…\n'
                           'I have changed my mind about…\n'
                           'I want people to feel ______ after talking with me.\n'
                           '\n'
                           'Look for patterns.\n'
                           '\n'
                           'SPEAK IT OUT\n'
                           '\n'
                           'Speak for two minutes:\n'
                           '\n'
                           '“What do you want to be known for professionally?”\n'
                           '\n'
                           'Do not describe a title.\n'
                           '\n'
                           'Describe a way of thinking and a way of affecting other people.\n'
                           '\n'
                           'Your goal is not to sound like anyone else.\n'
                           '\n'
                           'Your goal is to sound unmistakably like yourself—only clearer.'}]},
 {'book_id': 'talk-with-matthew',
  'title': 'คุยกับน้องแม็ท',
  'subtitle': 'Bedtime & Growing Up Stories',
  'author': 'Polly Chen',
  'content_type': 'Book',
  'category': 'Family',
  'description': 'เรื่องสั้นสำหรับคุยกับน้องแม็ทก่อนนอนและช่วยฝึกการดูแลตัวเอง',
  'cover_emoji': '🐰',
  'chapters': [{'chapter_id': 'matthew-bedtime',
                'chapter_title': 'น้องแม็ทเข้านอนให้ไว',
                'order': 1,
                'content': '\n'
                           '\n'
                           'Matthew, Teacher Gift said that Matthew should go to bed early.   Its time '
                           'for us to take a shower and brush our teeth now.\n'
                           '\n'
                           'Its time to drink your milk and get ready for bed. Once its bedtime, we '
                           'wont play anymore.\n'
                           '\n'
                           'Annalu also said that Matthew will be a good boy.\n'
                           '\n'
                           'Tonight, lets finish everything quickly, okay? Take a nice shower, brush '
                           'your teeth well, put on your pajamas, and get into bed.\n'
                           '\n'
                           'Before bedtime, we can choose one story, and Mommy will read it to '
                           'Matthew.\n'
                           '\n'
                           'When the story is finished, well turn off the light, close our eyes, and '
                           'let our bodies rest and grow strong.\n'
                           '\n'
                           'If Matthew goes to bed early, tomorrow morning youll wake up feeling fresh. '
                           'Youll have lots of energy to go to school, play with your friends, and '
                           'learn new things.\n'
                           '\n'
                           'Teacher Gift will be happy that Matthew gets enough rest.\n'
                           '\n'
                           'Mommy will be happy too, because Matthew is learning how to take care of '
                           'himself.\n'
                           '\n'
                           'Tonight, you dont have to fall asleep right away. Just lie still, hug your '
                           'pillow, breathe slowly, and relax.\n'
                           '\n'
                           'Good night, Matthew.\n'
                           '\n'
                           'You did a good job today.\n'
                           '\n'
                           'Tomorrow is another happy day.\n'
                           '\n'
                           '\n'
                           '\n'
                           'น้องแม็ทคะ คุณครูกิ๊ฟบอกว่า ให้น้องแม็ทนอนให้ไว '
                           'เราต้องไปอาบน้ำแปรงฟันแล้วนะคะ\n'
                           '\n'
                           'ถึงเวลากินนมนอนให้นอนเราจะไม่เล่นแล้วนะคะ\n'
                           '\n'
                           'Annalu also saids Matthew will be a good boy.\n'
                           '\n'
                           'คืนนี้น้องแม็ททำทุกอย่างให้เสร็จเร็ว ๆ นะคะ อาบน้ำให้สะอาด แปรงฟันให้สะอาด '
                           'ใส่ชุดนอน แล้วขึ้นเตียง\n'
                           '\n'
                           'ก่อนนอนเราสามารถเลือกนิทานหนึ่งเรื่อง แล้วคุณแม่จะอ่านให้น้องแม็ทฟัง\n'
                           '\n'
                           'พออ่านนิทานจบ เราจะปิดไฟ หลับตา และพักผ่อนให้ร่างกายแข็งแรง\n'
                           '\n'
                           'ถ้าน้องแม็ทนอนเร็ว พรุ่งนี้ตอนเช้าน้องแม็ทจะตื่นมาสดชื่น มีแรงไปโรงเรียน '
                           'มีแรงเล่นกับเพื่อน และมีแรงเรียนรู้สิ่งใหม่ ๆ\n'
                           '\n'
                           'คุณครูกิ๊ฟจะดีใจที่น้องแม็ทพักผ่อนเพียงพอ\n'
                           '\n'
                           'คุณแม่ก็จะดีใจเหมือนกัน เพราะน้องแม็ทกำลังเรียนรู้ที่จะดูแลตัวเอง\n'
                           '\n'
                           'คืนนี้เราไม่ต้องรีบนอนให้หลับทันทีนะคะ แค่นอนนิ่ง ๆ กอดหมอน หายใจช้า ๆ '
                           'แล้วพักผ่อน\n'
                           '\n'
                           'Good night, Matthew.\n'
                           '\n'
                           'You did a good job today.\n'
                           '\n'
                           'Tomorrow is another happy day.'},
               {'chapter_id': 'matthew-can-do-it',
                'chapter_title': 'น้องแม็ททำเองได้',
                'order': 2,
                'content': 'น้องแม็ทโตขึ้นทุกวันแล้วนะคะ และมีหลายอย่างที่น้องแม็ทสามารถทำเองได้\n'
                           '\n'
                           'ตอนเช้า เมื่อตื่นขึ้นมา น้องแม็ทลุกจากเตียง เก็บหมอน แล้วเดินไปล้างหน้า\n'
                           '\n'
                           'จากนั้นน้องแม็ทแปรงฟันและเตรียมตัวไปโรงเรียน\n'
                           '\n'
                           'ถ้ามีบางอย่างที่ยังทำไม่ได้ น้องแม็ทสามารถพูดว่า\n'
                           '\n'
                           'Mommy, can you help me please?\n'
                           '\n'
                           'การขอความช่วยเหลือไม่ใช่เรื่องน่าอายนะคะ '
                           'เด็กเก่งไม่จำเป็นต้องทำทุกอย่างได้ตั้งแต่ครั้งแรก\n'
                           '\n'
                           'เด็กเก่งคือเด็กที่ลองทำก่อน ถ้าทำไม่ได้ก็ถาม '
                           'แล้วเรียนรู้ว่าจะทำอย่างไรในครั้งต่อไป\n'
                           '\n'
                           'ที่โรงเรียน ถ้าน้องแม็ทไม่เข้าใจอะไร น้องแม็ทสามารถถามคุณครูได้\n'
                           '\n'
                           'ถ้าอยากเล่นกับเพื่อน น้องแม็ทสามารถพูดว่า\n'
                           '\n'
                           'Can I play with you?\n'
                           '\n'
                           'ถ้าเพื่อนกำลังใช้ของเล่นอยู่ น้องแม็ทสามารถรอ หรือเลือกเล่นอย่างอื่นก่อน\n'
                           '\n'
                           'ถ้าเกิดทำผิดพลาด น้องแม็ทพูดว่า\n'
                           '\n'
                           'Im sorry. I will try again.\n'
                           '\n'
                           'น้องแม็ทไม่จำเป็นต้องทำทุกอย่างให้สมบูรณ์แบบ\n'
                           '\n'
                           'สิ่งสำคัญคือ ลองทำ เรียนรู้ และลองใหม่\n'
                           '\n'
                           'ทุกครั้งที่น้องแม็ททำอะไรด้วยตัวเองได้อีกหนึ่งอย่าง '
                           'น้องแม็ทก็เก่งขึ้นอีกนิดหนึ่ง\n'
                           '\n'
                           'คุณแม่ภูมิใจเวลาน้องแม็ทพยายาม ไม่ใช่เฉพาะเวลาน้องแม็ททำสำเร็จ\n'
                           '\n'
                           'Tomorrow, lets try again.\n'
                           '\n'
                           'Matthew can learn.\n'
                           '\n'
                           'Matthew can ask.\n'
                           '\n'
                           'Matthew can try.\n'
                           '\n'
                           'And Matthew can do more and more by himself.'}],
  'audience': 'Kids'},
 {'book_id': 'milo-and-the-little-blue-star',
  'title': 'Milo and the Little Blue Star',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '⭐',
  'chapters': [{'chapter_id': 'milo-and-the-little-blue-star',
                'chapter_title': 'Milo and the Little Blue Star',
                'order': 1,
                'content': 'Once upon a time, there was a boy named Milo.\n'
                           '\n'
                           'Milo lived in a small house on a quiet hill.\n'
                           '\n'
                           'Every night, he liked to sit by his window and look at the stars.\n'
                           '\n'
                           'Some stars were big.\n'
                           '\n'
                           'Some stars were small.\n'
                           '\n'
                           'Some were bright.\n'
                           '\n'
                           'Some were very far away.\n'
                           '\n'
                           'But there was one little blue star that Milo loved most.\n'
                           '\n'
                           '“Good night, little star,” Milo said every night.\n'
                           '\n'
                           'And every night, the little blue star seemed to shine back at him.\n'
                           '\n'
                           'One evening, something strange happened.\n'
                           '\n'
                           'The little blue star began to fall.\n'
                           '\n'
                           'Down, down, down it came.\n'
                           '\n'
                           'Milo ran outside.\n'
                           '\n'
                           'A soft blue light landed in the grass behind his house.\n'
                           '\n'
                           'Milo walked closer.\n'
                           '\n'
                           'There, sitting in the grass, was a tiny girl.\n'
                           '\n'
                           'She had silver hair and a blue coat.\n'
                           '\n'
                           '“Hello,” she said.\n'
                           '\n'
                           'Milo blinked.\n'
                           '\n'
                           '“Are you… the little blue star?”\n'
                           '\n'
                           'The girl smiled.\n'
                           '\n'
                           '“My name is Lumi. I live on the little blue star.”\n'
                           '\n'
                           'Milo sat beside her.\n'
                           '\n'
                           '“Why did you come here?”\n'
                           '\n'
                           'Lumi looked up at the sky.\n'
                           '\n'
                           '“I lost something.”\n'
                           '\n'
                           '“What did you lose?”\n'
                           '\n'
                           "“I don't know.”\n"
                           '\n'
                           'Milo looked confused.\n'
                           '\n'
                           "“You lost something, but you don't know what it is?”\n"
                           '\n'
                           'Lumi nodded.\n'
                           '\n'
                           '“I only know that my star does not feel like home anymore.”\n'
                           '\n'
                           'Milo thought for a moment.\n'
                           '\n'
                           '“Maybe we can find it together.”\n'
                           '\n'
                           'So Milo and Lumi began their journey.\n'
                           '\n'
                           'First, they walked to the top of the hill.\n'
                           '\n'
                           'There they met an old man building a very tall fence.\n'
                           '\n'
                           '“What are you doing?” Milo asked.\n'
                           '\n'
                           '“I am making the highest fence in the world,” said the man.\n'
                           '\n'
                           '“Why?”\n'
                           '\n'
                           '“So nobody can come inside.”\n'
                           '\n'
                           'Lumi looked at the man.\n'
                           '\n'
                           '“Does that make you happy?”\n'
                           '\n'
                           'The old man stopped.\n'
                           '\n'
                           'He looked at his fence.\n'
                           '\n'
                           'Then he looked at the empty chair beside him.\n'
                           '\n'
                           '“No,” he said quietly.\n'
                           '\n'
                           '“But nobody can hurt me here.”\n'
                           '\n'
                           'Milo and Lumi walked away.\n'
                           '\n'
                           'Lumi said, “Maybe being safe is not the same as being happy.”\n'
                           '\n'
                           'Milo nodded.\n'
                           '\n'
                           'Next, they came to a beautiful garden.\n'
                           '\n'
                           'A woman was watering hundreds of flowers.\n'
                           '\n'
                           'Red flowers.\n'
                           '\n'
                           'Yellow flowers.\n'
                           '\n'
                           'Purple flowers.\n'
                           '\n'
                           'White flowers.\n'
                           '\n'
                           '“They are beautiful,” Lumi said.\n'
                           '\n'
                           'The woman smiled proudly.\n'
                           '\n'
                           '“I have more flowers than anyone in this town.”\n'
                           '\n'
                           '“Which one is your favorite?” Milo asked.\n'
                           '\n'
                           'The woman looked around.\n'
                           '\n'
                           '“My favorite?”\n'
                           '\n'
                           '“Yes.”\n'
                           '\n'
                           'She became quiet.\n'
                           '\n'
                           "“I don't know. I don't have time to know them one by one.”\n"
                           '\n'
                           'Lumi touched a small yellow flower.\n'
                           '\n'
                           '“This one smells sweet.”\n'
                           '\n'
                           'The woman stopped watering.\n'
                           '\n'
                           'She bent down and smelled it.\n'
                           '\n'
                           'For the first time that day, she smiled.\n'
                           '\n'
                           'Milo and Lumi continued walking.\n'
                           '\n'
                           'Lumi said, “Maybe having many things is not the same as loving them.”\n'
                           '\n'
                           'Milo nodded again.\n'
                           '\n'
                           'Later, they reached a small bridge.\n'
                           '\n'
                           'A little brown dog was sitting there alone.\n'
                           '\n'
                           'The dog looked cold.\n'
                           '\n'
                           'Milo took off his scarf and placed it around the dog.\n'
                           '\n'
                           '“Come with us,” he said.\n'
                           '\n'
                           'The dog wagged his tail.\n'
                           '\n'
                           'They called him Pip.\n'
                           '\n'
                           'Now there were three travelers.\n'
                           '\n'
                           'Milo, Lumi, and Pip.\n'
                           '\n'
                           'They walked until the sky became dark.\n'
                           '\n'
                           '“I am tired,” Lumi said.\n'
                           '\n'
                           'They sat under a large tree.\n'
                           '\n'
                           'Milo shared his bread with Lumi.\n'
                           '\n'
                           'He gave a small piece to Pip too.\n'
                           '\n'
                           'They did not have much.\n'
                           '\n'
                           'But they shared everything.\n'
                           '\n'
                           'Lumi looked at Milo.\n'
                           '\n'
                           '“Why did you help me?”\n'
                           '\n'
                           '“Because you needed help.”\n'
                           '\n'
                           '“Why did you help Pip?”\n'
                           '\n'
                           '“Because he was cold.”\n'
                           '\n'
                           '“But what do you get?”\n'
                           '\n'
                           'Milo laughed.\n'
                           '\n'
                           "“I don't know.”\n"
                           '\n'
                           'Lumi looked at Pip sleeping beside them.\n'
                           '\n'
                           'Then she smiled.\n'
                           '\n'
                           '“I think I understand something.”\n'
                           '\n'
                           '“What?”\n'
                           '\n'
                           '“When we care for someone, our heart becomes bigger.”\n'
                           '\n'
                           'The next morning, Lumi suddenly stopped.\n'
                           '\n'
                           '“My star!”\n'
                           '\n'
                           'Far above them, the little blue star was becoming dark.\n'
                           '\n'
                           '“I have to go home.”\n'
                           '\n'
                           'Milo felt a heavy feeling inside.\n'
                           '\n'
                           'He did not want Lumi to leave.\n'
                           '\n'
                           'But he knew she had to go.\n'
                           '\n'
                           '“How will you get back?”\n'
                           '\n'
                           'Lumi looked at Milo.\n'
                           '\n'
                           '“I think I know now.”\n'
                           '\n'
                           'She closed her eyes.\n'
                           '\n'
                           'The little blue star began to glow.\n'
                           '\n'
                           'A soft blue light came down from the sky.\n'
                           '\n'
                           'Before Lumi left, she hugged Milo.\n'
                           '\n'
                           '“I know what I lost.”\n'
                           '\n'
                           '“What was it?”\n'
                           '\n'
                           '“I thought home was a place.”\n'
                           '\n'
                           'She looked at Milo and Pip.\n'
                           '\n'
                           '“But home is also where we care for someone.”\n'
                           '\n'
                           'Milo smiled, though his eyes were wet.\n'
                           '\n'
                           '“Will I ever see you again?”\n'
                           '\n'
                           'Lumi pointed to the sky.\n'
                           '\n'
                           '“Look for the little blue star.”\n'
                           '\n'
                           'Then she rose into the light.\n'
                           '\n'
                           'Higher.\n'
                           '\n'
                           'Higher.\n'
                           '\n'
                           'Higher.\n'
                           '\n'
                           'Until she disappeared.\n'
                           '\n'
                           'That night, Milo sat by his window again.\n'
                           '\n'
                           'Pip sat beside him.\n'
                           '\n'
                           'The little blue star was brighter than ever.\n'
                           '\n'
                           'Milo waved.\n'
                           '\n'
                           'The star blinked once.\n'
                           '\n'
                           'Then twice.\n'
                           '\n'
                           'Milo smiled.\n'
                           '\n'
                           'From that day on, he remembered something important.\n'
                           '\n'
                           'A big house does not always make a home.\n'
                           '\n'
                           'Many things do not always make us rich.\n'
                           '\n'
                           'And being alone does not always make us safe.\n'
                           '\n'
                           'What makes life special is much simpler.\n'
                           '\n'
                           'It is the time we give.\n'
                           '\n'
                           'The care we share.\n'
                           '\n'
                           'And the people we choose to keep in our hearts.\n'
                           '\n'
                           'Milo looked at Pip.\n'
                           '\n'
                           '“Good night, little star.”\n'
                           '\n'
                           'Far away in the sky, Lumi smiled.\n'
                           '\n'
                           'And the little blue star shone brightly through the night.\n'
                           '\n'
                           'The End.'}],
  'audience': 'Kids'},
 {'book_id': 'prince-leo-and-the-valley-of-three-doors',
  'title': 'Prince Leo and the Valley of Three Doors',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '👑',
  'chapters': [{'chapter_id': 'prince-leo-and-the-valley-of-three-doors',
                'chapter_title': 'Prince Leo and the Valley of Three Doors',
                'order': 1,
                'content': 'Prince Leo and the Valley of Three Doors\n'
                           'Once upon a time, there was a young prince named Leo.\n'
                           '\n'
                           'Prince Leo lived in a beautiful castle near the sea.\n'
                           '\n'
                           'He had a soft bed.\n'
                           '\n'
                           'He had delicious food.\n'
                           '\n'
                           'He had teachers, guards, horses, and many toys.\n'
                           '\n'
                           'But Leo had one problem.\n'
                           '\n'
                           'He was afraid of making mistakes.\n'
                           '\n'
                           'Every time he had to make a choice, he asked someone else.\n'
                           '\n'
                           '“What should I wear?”\n'
                           '\n'
                           '“What should I say?”\n'
                           '\n'
                           '“Which road should I take?”\n'
                           '\n'
                           '“What if I choose the wrong one?”\n'
                           '\n'
                           'One morning, the king called Leo to the throne room.\n'
                           '\n'
                           '“My son,” said the king, “one day you will lead this kingdom.”\n'
                           '\n'
                           'Leo looked worried.\n'
                           '\n'
                           '“But what if I make a mistake?”\n'
                           '\n'
                           'The king smiled.\n'
                           '\n'
                           '“You will.”\n'
                           '\n'
                           "Leo's eyes became wide.\n"
                           '\n'
                           '“I will?”\n'
                           '\n'
                           '“Of course.”\n'
                           '\n'
                           'The king laughed.\n'
                           '\n'
                           '“Everyone makes mistakes. A good leader is not someone who never makes '
                           'mistakes.”\n'
                           '\n'
                           '“Then what is a good leader?”\n'
                           '\n'
                           'The king handed Leo a small golden compass.\n'
                           '\n'
                           '“Go to the Valley of Three Doors. Find the answer yourself.”\n'
                           '\n'
                           'Leo did not want to go.\n'
                           '\n'
                           'But for the first time, he decided not to ask anyone what he should do.\n'
                           '\n'
                           'He packed some bread and water.\n'
                           '\n'
                           'Then he climbed onto his horse, Sunny.\n'
                           '\n'
                           'Together, they left the castle.\n'
                           '\n'
                           'They crossed a green forest.\n'
                           '\n'
                           'They crossed a shallow river.\n'
                           '\n'
                           'They climbed a rocky hill.\n'
                           '\n'
                           'Finally, they reached the Valley of Three Doors.\n'
                           '\n'
                           'There were three giant doors standing in the middle of the valley.\n'
                           '\n'
                           'The first door was made of gold.\n'
                           '\n'
                           'The second door was made of stone.\n'
                           '\n'
                           'The third door was small and made of old wood.\n'
                           '\n'
                           'A sign said:\n'
                           '\n'
                           'CHOOSE ONE.\n'
                           '\n'
                           'Leo looked at the doors.\n'
                           '\n'
                           '“The golden door looks important,” he said.\n'
                           '\n'
                           '“But maybe that is too easy.”\n'
                           '\n'
                           'He looked at the stone door.\n'
                           '\n'
                           '“That one looks strong.”\n'
                           '\n'
                           'Then he looked at the wooden door.\n'
                           '\n'
                           '“It looks boring.”\n'
                           '\n'
                           'Leo waited.\n'
                           '\n'
                           'He wanted someone to tell him which door was correct.\n'
                           '\n'
                           'But nobody came.\n'
                           '\n'
                           'Sunny ate some grass.\n'
                           '\n'
                           '“Very helpful,” Leo said.\n'
                           '\n'
                           'Finally, Leo took a deep breath.\n'
                           '\n'
                           '“I will choose the golden door.”\n'
                           '\n'
                           'He opened it.\n'
                           '\n'
                           'Behind the door was a huge room filled with treasure.\n'
                           '\n'
                           'Gold.\n'
                           '\n'
                           'Jewels.\n'
                           '\n'
                           'Crowns.\n'
                           '\n'
                           'Silver cups.\n'
                           '\n'
                           'Leo smiled.\n'
                           '\n'
                           '“I chose correctly!”\n'
                           '\n'
                           'Then the golden door closed behind him.\n'
                           '\n'
                           'Suddenly, the room became very dark.\n'
                           '\n'
                           'Leo heard a voice.\n'
                           '\n'
                           '“You may take everything.”\n'
                           '\n'
                           'A large treasure box opened.\n'
                           '\n'
                           '“But you must leave your horse behind.”\n'
                           '\n'
                           'Leo looked through a small window.\n'
                           '\n'
                           'Sunny was waiting outside.\n'
                           '\n'
                           'Leo thought about the treasure.\n'
                           '\n'
                           'Then he shook his head.\n'
                           '\n'
                           '“No.”\n'
                           '\n'
                           'The voice asked, “Why?”\n'
                           '\n'
                           '“Because Sunny is my friend.”\n'
                           '\n'
                           'The treasure disappeared.\n'
                           '\n'
                           'The golden door opened.\n'
                           '\n'
                           'Leo walked out.\n'
                           '\n'
                           'He had learned his first lesson.\n'
                           '\n'
                           'Not everything valuable shines.\n'
                           '\n'
                           'Next, Leo opened the stone door.\n'
                           '\n'
                           'Behind it was a village.\n'
                           '\n'
                           'The people were running everywhere.\n'
                           '\n'
                           'A storm had broken their bridge.\n'
                           '\n'
                           'Nobody could cross the river.\n'
                           '\n'
                           'A woman cried, “My children are on the other side!”\n'
                           '\n'
                           'Leo wanted to help.\n'
                           '\n'
                           'But he had never built a bridge before.\n'
                           '\n'
                           "“I don't know how,” he said.\n"
                           '\n'
                           'An old builder looked at him.\n'
                           '\n'
                           '“Neither do I, not alone.”\n'
                           '\n'
                           'Leo was surprised.\n'
                           '\n'
                           '“You are a builder.”\n'
                           '\n'
                           '“Yes,” said the man. “But this bridge is too large for one person.”\n'
                           '\n'
                           'Leo looked around.\n'
                           '\n'
                           'There were farmers.\n'
                           '\n'
                           'Carpenters.\n'
                           '\n'
                           'Fishermen.\n'
                           '\n'
                           'Children.\n'
                           '\n'
                           'Strong people and clever people.\n'
                           '\n'
                           'Then Leo had an idea.\n'
                           '\n'
                           '“You!” he said to the carpenters. “Find strong wood.”\n'
                           '\n'
                           '“You!” he said to the fishermen. “Bring ropes.”\n'
                           '\n'
                           '“And everyone else, help carry the stones.”\n'
                           '\n'
                           'The villagers worked together.\n'
                           '\n'
                           'Leo did not know how to do every job.\n'
                           '\n'
                           'But he listened.\n'
                           '\n'
                           'He asked questions.\n'
                           '\n'
                           'He helped people work together.\n'
                           '\n'
                           'By sunset, the bridge was ready.\n'
                           '\n'
                           'The woman crossed the river and hugged her children.\n'
                           '\n'
                           'The villagers cheered.\n'
                           '\n'
                           'Leo smiled.\n'
                           '\n'
                           'He had learned his second lesson.\n'
                           '\n'
                           'A leader does not need to know everything.\n'
                           '\n'
                           'A leader helps people work together.\n'
                           '\n'
                           'Then Leo returned to the valley.\n'
                           '\n'
                           'Only the small wooden door remained.\n'
                           '\n'
                           'Leo opened it.\n'
                           '\n'
                           'Behind it was a dark forest.\n'
                           '\n'
                           '“No treasure?”\n'
                           '\n'
                           'No answer.\n'
                           '\n'
                           '“No village?”\n'
                           '\n'
                           'Nothing.\n'
                           '\n'
                           'Leo entered the forest.\n'
                           '\n'
                           'Soon, he heard a cry.\n'
                           '\n'
                           '“Help!”\n'
                           '\n'
                           'Leo followed the sound.\n'
                           '\n'
                           'A young fox was trapped under a fallen branch.\n'
                           '\n'
                           'Leo tried to lift the branch.\n'
                           '\n'
                           'It was too heavy.\n'
                           '\n'
                           'He pushed again.\n'
                           '\n'
                           'Nothing happened.\n'
                           '\n'
                           'Leo felt tired.\n'
                           '\n'
                           "“I can't do it.”\n"
                           '\n'
                           'The fox looked at him.\n'
                           '\n'
                           'Leo almost walked away.\n'
                           '\n'
                           'Then he remembered the village.\n'
                           '\n'
                           'He looked around carefully.\n'
                           '\n'
                           'He found a long piece of wood.\n'
                           '\n'
                           'He placed it under the branch like a lever.\n'
                           '\n'
                           'Then he pushed.\n'
                           '\n'
                           'The branch moved.\n'
                           '\n'
                           'The fox escaped.\n'
                           '\n'
                           '“You did it!” said the fox.\n'
                           '\n'
                           'Leo smiled.\n'
                           '\n'
                           '“Not the first way.”\n'
                           '\n'
                           'They walked together through the forest.\n'
                           '\n'
                           'Soon they reached a deep hole in the road.\n'
                           '\n'
                           'Leo could not cross.\n'
                           '\n'
                           'The fox said, “Follow me.”\n'
                           '\n'
                           'The fox showed Leo a narrow path around the hole.\n'
                           '\n'
                           'Leo laughed.\n'
                           '\n'
                           '“This time, you helped me.”\n'
                           '\n'
                           'The fox smiled.\n'
                           '\n'
                           '“No one is strong all the time.”\n'
                           '\n'
                           'At the end of the forest, Leo found a small mirror.\n'
                           '\n'
                           'There were no jewels around it.\n'
                           '\n'
                           'No magic.\n'
                           '\n'
                           'Only a simple mirror.\n'
                           '\n'
                           'Under it were the words:\n'
                           '\n'
                           'THE LAST DOOR SHOWS THE LEADER.\n'
                           '\n'
                           'Leo looked into the mirror.\n'
                           '\n'
                           'He only saw himself.\n'
                           '\n'
                           '“At first, I thought a leader had to be perfect,” Leo said.\n'
                           '\n'
                           'The fox sat beside him.\n'
                           '\n'
                           '“And now?”\n'
                           '\n'
                           'Leo thought for a moment.\n'
                           '\n'
                           '“A leader has to choose, even when he is not sure.”\n'
                           '\n'
                           'The fox nodded.\n'
                           '\n'
                           '“A leader has to listen to people who know more.”\n'
                           '\n'
                           'The fox nodded again.\n'
                           '\n'
                           '“And when something does not work, a leader has to try another way.”\n'
                           '\n'
                           'The fox smiled.\n'
                           '\n'
                           'Leo finally understood.\n'
                           '\n'
                           'The golden compass in his pocket began to shine.\n'
                           '\n'
                           'It pointed toward home.\n'
                           '\n'
                           'When Leo returned to the castle, the king was waiting.\n'
                           '\n'
                           '“Well?” the king asked.\n'
                           '\n'
                           '“Did you find the answer?”\n'
                           '\n'
                           'Leo gave the golden compass back to his father.\n'
                           '\n'
                           '“Yes.”\n'
                           '\n'
                           '“And what makes a good leader?”\n'
                           '\n'
                           'Leo smiled.\n'
                           '\n'
                           '“Someone who cares about what matters.”\n'
                           '\n'
                           'The king listened.\n'
                           '\n'
                           '“Someone who brings people together.”\n'
                           '\n'
                           'The king nodded.\n'
                           '\n'
                           '“And someone who keeps thinking when the first answer does not work.”\n'
                           '\n'
                           'The king smiled.\n'
                           '\n'
                           '“Anything else?”\n'
                           '\n'
                           'Leo laughed.\n'
                           '\n'
                           '“Yes.”\n'
                           '\n'
                           '“A good leader will still make mistakes.”\n'
                           '\n'
                           'The king laughed too.\n'
                           '\n'
                           '“Now you are ready to learn.”\n'
                           '\n'
                           'From that day on, Leo still made mistakes.\n'
                           '\n'
                           'Sometimes he chose the wrong road.\n'
                           '\n'
                           'Sometimes his ideas did not work.\n'
                           '\n'
                           'Sometimes he needed help.\n'
                           '\n'
                           'But he was no longer afraid of those things.\n'
                           '\n'
                           'Because he knew that courage was not about always knowing the answer.\n'
                           '\n'
                           'Courage was being willing to find it.\n'
                           '\n'
                           'And many years later, when Leo became king, people did not remember him '
                           'because he had the biggest castle or the most treasure.\n'
                           '\n'
                           'They remembered him because when people had a problem, King Leo listened.\n'
                           '\n'
                           'When people had different ideas, he brought them together.\n'
                           '\n'
                           'And when the kingdom faced a difficult road, he did not simply ask,\n'
                           '\n'
                           '“Who made the mistake?”\n'
                           '\n'
                           'He asked,\n'
                           '\n'
                           '“What can we learn, and what should we do next?”\n'
                           '\n'
                           'And that was why the people trusted him.\n'
                           '\n'
                           'The End.'}],
  'audience': 'Kids'},
 {'book_id': 'the-little-prince-simple-english-retelling',
  'title': 'The Little Prince',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🌟',
  'chapters': [{'chapter_id': 'the-little-prince-simple-english-retelling',
                'chapter_title': 'The Little Prince',
                'order': 1,
                'content': 'The story begins with a pilot.\n'
                           '\n'
                           'When he was a little boy, he liked to draw. One day, he drew a picture of a '
                           'snake that had eaten an elephant. But the adults thought his picture was '
                           'only a hat. They did not understand him.\n'
                           '\n'
                           'So he stopped drawing and grew up to become a pilot.\n'
                           '\n'
                           'Many years later, his airplane crashed in the Sahara Desert. He was alone '
                           'and had very little water.\n'
                           '\n'
                           'Then, suddenly, he heard a small voice.\n'
                           '\n'
                           '“Please draw me a sheep.”\n'
                           '\n'
                           'The pilot was surprised. In the middle of the desert stood a small boy with '
                           'golden hair.\n'
                           '\n'
                           'This boy was the Little Prince.\n'
                           '\n'
                           'The Little Prince came from a very small planet called asteroid B-612. His '
                           'planet was so small that he could watch many sunsets simply by moving his '
                           'chair.\n'
                           '\n'
                           'On his planet, the Little Prince had three volcanoes and a beautiful rose.\n'
                           '\n'
                           'The rose was very special to him. She was beautiful, proud, and sometimes '
                           'difficult. She often asked him to take care of her.\n'
                           '\n'
                           'The Little Prince loved her, but he was too young to understand his '
                           'feelings.\n'
                           '\n'
                           'He became confused and decided to leave his planet.\n'
                           '\n'
                           'He traveled from one small planet to another.\n'
                           '\n'
                           'On the first planet, he met a king who wanted to rule everything.\n'
                           '\n'
                           'On another planet, he met a man who wanted everyone to admire him.\n'
                           '\n'
                           'Then he met a businessman who spent all his time counting stars because he '
                           'believed he owned them.\n'
                           '\n'
                           'The Little Prince thought these adults were very strange.\n'
                           '\n'
                           'They were always busy with things they believed were important, but they '
                           'often forgot how to enjoy life.\n'
                           '\n'
                           'Finally, the Little Prince arrived on Earth.\n'
                           '\n'
                           'There, he saw a garden full of roses.\n'
                           '\n'
                           'He became very sad.\n'
                           '\n'
                           'He had always believed that his rose was the only rose in the universe. Now '
                           'he saw thousands of roses that looked just like her.\n'
                           '\n'
                           'Then he met a fox.\n'
                           '\n'
                           'The fox asked the Little Prince to become his friend.\n'
                           '\n'
                           'At first, the Little Prince did not understand.\n'
                           '\n'
                           'The fox explained that when two people spend time together and care about '
                           'each other, they create a special bond.\n'
                           '\n'
                           'Slowly, the Little Prince understood.\n'
                           '\n'
                           'His rose was special not because she was the only rose in the world.\n'
                           '\n'
                           'She was special because he had cared for her.\n'
                           '\n'
                           'He had watered her.\n'
                           '\n'
                           'He had protected her.\n'
                           '\n'
                           'He had listened to her.\n'
                           '\n'
                           'And he had spent his time with her.\n'
                           '\n'
                           'That was what made her important.\n'
                           '\n'
                           'The fox also taught him that the most important things in life cannot '
                           'always be seen with our eyes.\n'
                           '\n'
                           'The Little Prince began to understand that he still loved his rose.\n'
                           '\n'
                           'He wanted to return home.\n'
                           '\n'
                           'Meanwhile, the pilot was trying to repair his airplane.\n'
                           '\n'
                           'The pilot and the Little Prince became close friends.\n'
                           '\n'
                           'But one day, the Little Prince decided that he had to return to his '
                           'planet.\n'
                           '\n'
                           'Leaving was difficult.\n'
                           '\n'
                           'The pilot did not want to lose his friend.\n'
                           '\n'
                           'But the Little Prince told him to look at the stars.\n'
                           '\n'
                           'Whenever the pilot looked at them, he could remember the Little Prince and '
                           'imagine him laughing somewhere far away.\n'
                           '\n'
                           'In the end, the pilot was left alone in the desert.\n'
                           '\n'
                           'But he never forgot his little friend.\n'
                           '\n'
                           'And whenever he looked at the stars, they seemed different.\n'
                           '\n'
                           'Because somewhere among them was a small planet, a rose, and a little '
                           'prince he loved.\n'
                           '\n'
                           'The story reminds us that love is not about finding something perfect.\n'
                           '\n'
                           'It is about the time, care, and love we give to someone.\n'
                           '\n'
                           'And sometimes the most important things in life are the things we cannot '
                           'see.'}],
  'audience': 'Kids'},
 {'book_id': 'prince-pipo',
  'title': 'Prince Pipo',
  'subtitle': "Children's Story",
  'author': 'Pierre Gripari',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🐴',
  'chapters': [{'chapter_id': 'prince-pipo',
                'chapter_title': 'Prince Pipo',
                'order': 1,
                'content': '### *Prince Pipo* is a story by French writer Pierre Gripari. Pipo is about '
                           'fifteen years old when he begins his journey into the world. Official '
                           'descriptions of the book present it as a story about growing up, freedom, '
                           'adventure, and discovering that parents and the adult world are not always '
                           'what a child imagined. \n'
                           '\n'
                           'Before Pipo is born, his father wants very much to have a son.\n'
                           '\n'
                           'He goes to a powerful witch and asks for help.\n'
                           '\n'
                           'There, he finds the child who will become Pipo.\n'
                           '\n'
                           'But there is a condition.\n'
                           '\n'
                           'Pipo must agree to become his son.\n'
                           '\n'
                           'To persuade him, the future father promises Pipo something very important: '
                           'when Pipo is fifteen years old, he will receive a beautiful red horse. \n'
                           '\n'
                           'Pipo grows up believing that he is a prince.\n'
                           '\n'
                           'He lives in a world that feels safe and wonderful.\n'
                           '\n'
                           'But when he becomes older, things begin to change.\n'
                           '\n'
                           'He discovers that the world is not as simple as he thought.\n'
                           '\n'
                           'His parents are not perfect.\n'
                           '\n'
                           'Life is not always fair.\n'
                           '\n'
                           'And promises are not always kept.\n'
                           '\n'
                           'One day, Pipo learns something painful about his family and his life.\n'
                           '\n'
                           'The world of his childhood seems to disappear.\n'
                           '\n'
                           'So Pipo leaves home.\n'
                           '\n'
                           'He begins a long journey with his horse, who is also called Pipo.\n'
                           '\n'
                           'The horse is much more than an animal.\n'
                           '\n'
                           "He becomes Pipo's companion as the boy travels into an unfamiliar world.\n"
                           '\n'
                           'During the journey, Pipo meets danger again and again.\n'
                           '\n'
                           'He faces witches.\n'
                           '\n'
                           'He faces a volcano.\n'
                           '\n'
                           'He meets violence, fear, loneliness, and the strange problems of the adult '
                           'world. The story also includes magical transformations: at one point Pipo '
                           'becomes a dragon. \n'
                           '\n'
                           'Pipo wants freedom.\n'
                           '\n'
                           'But he discovers that freedom is not simply doing whatever you want.\n'
                           '\n'
                           'Freedom can also mean making difficult choices and accepting the results.\n'
                           '\n'
                           'As he travels, Pipo learns to survive without the protection he had as a '
                           'child.\n'
                           '\n'
                           'He meets people who help him and people who hurt him.\n'
                           '\n'
                           'Sometimes he is brave.\n'
                           '\n'
                           'Sometimes he is confused.\n'
                           '\n'
                           'Sometimes he makes mistakes.\n'
                           '\n'
                           'But every experience changes him.\n'
                           '\n'
                           'His journey is not only about finding a kingdom.\n'
                           '\n'
                           'It is also about understanding himself.\n'
                           '\n'
                           'Eventually, Pipo meets Princess Popi.\n'
                           '\n'
                           'Her arrival changes his story.\n'
                           '\n'
                           'For the first time, Pipo begins to understand another kind of love—a love '
                           'that is different from the love between a child and his parents.\n'
                           '\n'
                           "Popi becomes very important to him, and Pipo's adventures begin to lead him "
                           'toward a new stage of life. \n'
                           '\n'
                           'By the end, Pipo is no longer the same boy who left home.\n'
                           '\n'
                           'He has seen that adults are imperfect.\n'
                           '\n'
                           'He has learned that the world can be frightening.\n'
                           '\n'
                           'He has learned about loneliness, courage, love, and responsibility.\n'
                           '\n'
                           "In many ways, Pipo's journey is the journey every child makes while growing "
                           'up.\n'
                           '\n'
                           'When we are small, our parents can seem like kings and queens.\n'
                           '\n'
                           'Our home can feel like the whole world.\n'
                           '\n'
                           'Then one day, we discover that our parents are ordinary people too.\n'
                           '\n'
                           'They can make mistakes.\n'
                           '\n'
                           'They cannot protect us from everything.\n'
                           '\n'
                           'And we must slowly learn to make our own way in the world.\n'
                           '\n'
                           'That is what happens to Prince Pipo.\n'
                           '\n'
                           'His adventures with witches, dragons, danger, his horse, and Princess Popi '
                           'are magical.\n'
                           '\n'
                           'But underneath the magic is a very human story:\n'
                           '\n'
                           'a child becoming an adult.'}],
  'audience': 'Kids'},
 {'book_id': 'the-two-lanterns',
  'title': 'The Two Lanterns',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🏮',
  'chapters': [{'chapter_id': 'the-two-lanterns',
                'chapter_title': 'The Two Lanterns',
                'order': 1,
                'content': 'Once upon a time, there were two children named Kathy and Kathin.\n'
                           '\n'
                           'Kathy was the older sister.\n'
                           '\n'
                           'Kathin was the younger brother.\n'
                           '\n'
                           'They loved each other very much.\n'
                           '\n'
                           'But they also fought.\n'
                           '\n'
                           'A lot.\n'
                           '\n'
                           'They fought about toys.\n'
                           '\n'
                           'They fought about books.\n'
                           '\n'
                           'They fought about who sat next to Mom.\n'
                           '\n'
                           'They even fought about who got the bigger piece of banana.\n'
                           '\n'
                           'One morning, Kathy was building a tall tower with colorful blocks.\n'
                           '\n'
                           'She worked very carefully.\n'
                           '\n'
                           'One block.\n'
                           '\n'
                           'Two blocks.\n'
                           '\n'
                           'Three blocks.\n'
                           '\n'
                           'Higher and higher.\n'
                           '\n'
                           'Kathin walked into the room.\n'
                           '\n'
                           '“I want the blue block,” he said.\n'
                           '\n'
                           '“I need it for my tower,” said Kathy.\n'
                           '\n'
                           '“But I want it!”\n'
                           '\n'
                           '“You can have another one.”\n'
                           '\n'
                           '“No! I want that one!”\n'
                           '\n'
                           'Kathin reached for the blue block.\n'
                           '\n'
                           '“Stop!” Kathy shouted.\n'
                           '\n'
                           'Kathin pulled.\n'
                           '\n'
                           'Kathy pulled back.\n'
                           '\n'
                           'Then—\n'
                           '\n'
                           'CRASH!\n'
                           '\n'
                           'The whole tower fell down.\n'
                           '\n'
                           'Kathy stared at the blocks on the floor.\n'
                           '\n'
                           'Then she looked at Kathin.\n'
                           '\n'
                           '“You broke it!”\n'
                           '\n'
                           '“You didn’t give me the block!”\n'
                           '\n'
                           '“You always ruin everything!”\n'
                           '\n'
                           '“You never share!”\n'
                           '\n'
                           'They both shouted.\n'
                           '\n'
                           'Mom came into the room.\n'
                           '\n'
                           'She looked at Kathy.\n'
                           '\n'
                           'She looked at Kathin.\n'
                           '\n'
                           'Then she looked at the blocks all over the floor.\n'
                           '\n'
                           'Mom did not ask, “Who started it?”\n'
                           '\n'
                           'Instead, she said,\n'
                           '\n'
                           '“I see two angry children.”\n'
                           '\n'
                           '“I’m not angry!” said Kathy.\n'
                           '\n'
                           '“Yes, you are!” said Kathin.\n'
                           '\n'
                           '“No, I’m not!”\n'
                           '\n'
                           '“Yes, you are!”\n'
                           '\n'
                           'Mom raised one eyebrow.\n'
                           '\n'
                           'Kathy and Kathin became quiet.\n'
                           '\n'
                           '“Let’s give our bodies a little time,” Mom said.\n'
                           '\n'
                           'Kathy sat on the sofa.\n'
                           '\n'
                           'Kathin sat on the floor.\n'
                           '\n'
                           'Neither of them looked at the other.\n'
                           '\n'
                           'A few minutes later, Mom said,\n'
                           '\n'
                           '“I have a story for you.”\n'
                           '\n'
                           '“We don’t want a story,” Kathy said.\n'
                           '\n'
                           '“Yes,” said Kathin.\n'
                           '\n'
                           'Mom smiled.\n'
                           '\n'
                           '“That’s okay. I will tell it to myself.”\n'
                           '\n'
                           'Kathy looked at Kathin.\n'
                           '\n'
                           'Kathin looked at Kathy.\n'
                           '\n'
                           'They both moved a little closer.\n'
                           '\n'
                           'Mom began.\n'
                           '\n'
                           '“Long ago, in a little village near the sea, there were two lanterns.”\n'
                           '\n'
                           '“One lantern was red.\n'
                           '\n'
                           'The other lantern was blue.\n'
                           '\n'
                           'They lived together in a small lighthouse.”\n'
                           '\n'
                           '“The red lantern was very bright.\n'
                           '\n'
                           '‘I can shine farther than you,’ said the red lantern.\n'
                           '\n'
                           '“The blue lantern was smaller.\n'
                           '\n'
                           'But its light was soft and clear.\n'
                           '\n'
                           '‘I help ships see the rocks,’ said the blue lantern.\n'
                           '\n'
                           '“I am more important,” said the red lantern.\n'
                           '\n'
                           '“No, I am more important,” said the blue lantern.\n'
                           '\n'
                           'Every night, they argued.\n'
                           '\n'
                           '“My light is bigger!”\n'
                           '\n'
                           '“My light is better!”\n'
                           '\n'
                           '“I shine first!”\n'
                           '\n'
                           '“No, I shine first!”\n'
                           '\n'
                           'One evening, a big storm came.\n'
                           '\n'
                           'The wind blew hard.\n'
                           '\n'
                           'The rain hit the windows.\n'
                           '\n'
                           'BOOM!\n'
                           '\n'
                           'Thunder shook the lighthouse.\n'
                           '\n'
                           'Far away, a small boat was trying to get home.\n'
                           '\n'
                           'Inside the boat was a fisherman.\n'
                           '\n'
                           'He could not see the shore.\n'
                           '\n'
                           'He could not see the rocks.\n'
                           '\n'
                           'He looked for the lighthouse.\n'
                           '\n'
                           'But inside the lighthouse, the two lanterns were still fighting.\n'
                           '\n'
                           '“I should stand in the front!” said the red lantern.\n'
                           '\n'
                           '“No! It is my turn!” said the blue lantern.\n'
                           '\n'
                           'They pushed against each other.\n'
                           '\n'
                           'The red lantern moved left.\n'
                           '\n'
                           'The blue lantern moved right.\n'
                           '\n'
                           'Suddenly—\n'
                           '\n'
                           'CLICK.\n'
                           '\n'
                           'Both lights went out.\n'
                           '\n'
                           '“Oh no,” said the red lantern.\n'
                           '\n'
                           '“Oh no,” said the blue lantern.\n'
                           '\n'
                           'Outside, the little boat moved closer to the rocks.\n'
                           '\n'
                           'The two lanterns became very quiet.\n'
                           '\n'
                           '“This happened because of you,” said the red lantern.\n'
                           '\n'
                           '“No, because of you,” said the blue lantern.\n'
                           '\n'
                           'They almost started fighting again.\n'
                           '\n'
                           'Then the blue lantern stopped.\n'
                           '\n'
                           '“Wait.”\n'
                           '\n'
                           'The red lantern stopped too.\n'
                           '\n'
                           '“What?”\n'
                           '\n'
                           '“The boat.”\n'
                           '\n'
                           'They both looked out the window.\n'
                           '\n'
                           'They could see the small boat in the storm.\n'
                           '\n'
                           'The fisherman needed them.\n'
                           '\n'
                           'For the first time, the two lanterns forgot about who was right.\n'
                           '\n'
                           '“We need to help,” said the red lantern.\n'
                           '\n'
                           '“But our lights are off,” said the blue lantern.\n'
                           '\n'
                           '“What can we do?”\n'
                           '\n'
                           'They thought.\n'
                           '\n'
                           'And thought.\n'
                           '\n'
                           'Then the blue lantern said,\n'
                           '\n'
                           '“You are taller than me.”\n'
                           '\n'
                           '“Yes.”\n'
                           '\n'
                           '“And I am small enough to reach the little switch behind you.”\n'
                           '\n'
                           'The red lantern looked surprised.\n'
                           '\n'
                           '“So?”\n'
                           '\n'
                           '“If you move a little, I can reach it.”\n'
                           '\n'
                           'The red lantern moved.\n'
                           '\n'
                           'The blue lantern stretched.\n'
                           '\n'
                           '“A little more,” said the blue lantern.\n'
                           '\n'
                           'The red lantern moved again.\n'
                           '\n'
                           '“Got it!”\n'
                           '\n'
                           'CLICK.\n'
                           '\n'
                           'The blue light came back on.\n'
                           '\n'
                           'A soft blue light shone across the sea.\n'
                           '\n'
                           '“I can see the rocks,” said the blue lantern.\n'
                           '\n'
                           '“But the boat is still too far away.”\n'
                           '\n'
                           'The red lantern had an idea.\n'
                           '\n'
                           '“Turn your light toward the rocks.”\n'
                           '\n'
                           '“Why?”\n'
                           '\n'
                           '“Trust me.”\n'
                           '\n'
                           'The blue lantern turned toward the rocks.\n'
                           '\n'
                           'The red lantern stood behind it.\n'
                           '\n'
                           'The blue light showed the dangerous rocks.\n'
                           '\n'
                           'The red light shone far across the sea.\n'
                           '\n'
                           'Together, the two lights made a clear path.\n'
                           '\n'
                           'The fisherman saw them.\n'
                           '\n'
                           '“Ah! The lighthouse!”\n'
                           '\n'
                           'He turned the boat.\n'
                           '\n'
                           'He moved away from the rocks.\n'
                           '\n'
                           'Slowly, safely, he came home.\n'
                           '\n'
                           'Inside the lighthouse, the red lantern and the blue lantern watched the '
                           'boat reach the village.\n'
                           '\n'
                           '“We did it,” said the red lantern.\n'
                           '\n'
                           '“Yes,” said the blue lantern.\n'
                           '\n'
                           'Then the red lantern became quiet.\n'
                           '\n'
                           '“I’m sorry I said I was more important.”\n'
                           '\n'
                           'The blue lantern looked down.\n'
                           '\n'
                           '“I’m sorry too.”\n'
                           '\n'
                           'The red lantern said,\n'
                           '\n'
                           '“I still like being bright.”\n'
                           '\n'
                           '“That’s okay,” said the blue lantern.\n'
                           '\n'
                           '“I still like being blue.”\n'
                           '\n'
                           '“That’s okay too.”\n'
                           '\n'
                           'They both laughed.\n'
                           '\n'
                           'The next night, the red lantern shone far across the sea.\n'
                           '\n'
                           'The blue lantern showed the rocks.\n'
                           '\n'
                           'Sometimes the red lantern went first.\n'
                           '\n'
                           'Sometimes the blue lantern went first.\n'
                           '\n'
                           'Sometimes they still argued.\n'
                           '\n'
                           '“I want that spot!”\n'
                           '\n'
                           '“But I was here first!”\n'
                           '\n'
                           '“You had it yesterday!”\n'
                           '\n'
                           'They were not perfect.\n'
                           '\n'
                           'But when they started getting very angry, one of them would say,\n'
                           '\n'
                           '“Boat.”\n'
                           '\n'
                           'The other would stop.\n'
                           '\n'
                           '“Boat?”\n'
                           '\n'
                           '“Yes. Remember the boat.”\n'
                           '\n'
                           'Then they would remember something important.\n'
                           '\n'
                           'They did not have to be the same.\n'
                           '\n'
                           'They did not have to agree about everything.\n'
                           '\n'
                           'They could both want something.\n'
                           '\n'
                           'They could both feel angry.\n'
                           '\n'
                           'But they did not have to hurt each other.\n'
                           '\n'
                           'And sometimes, when there was only one good place to stand, they learned to '
                           'say,\n'
                           '\n'
                           '“You go first this time.”\n'
                           '\n'
                           'Or,\n'
                           '\n'
                           '“Let’s take turns.”\n'
                           '\n'
                           'Or even,\n'
                           '\n'
                           '“I’m still angry, but I don’t want to fight.”\n'
                           '\n'
                           'And little by little, the lighthouse became a happier place.\n'
                           '\n'
                           'Mom stopped talking.\n'
                           '\n'
                           'Kathy and Kathin were quiet.\n'
                           '\n'
                           'Kathin looked at the blocks on the floor.\n'
                           '\n'
                           'Then he looked at Kathy.\n'
                           '\n'
                           '“I wanted the blue block.”\n'
                           '\n'
                           '“I know,” said Kathy.\n'
                           '\n'
                           '“I was still using it.”\n'
                           '\n'
                           '“I know.”\n'
                           '\n'
                           'Kathin touched one of the fallen blocks.\n'
                           '\n'
                           '“I’m sorry I pulled it.”\n'
                           '\n'
                           'Kathy was quiet for a moment.\n'
                           '\n'
                           'Then she said,\n'
                           '\n'
                           '“I’m sorry I said you ruin everything.”\n'
                           '\n'
                           'Kathin looked at her.\n'
                           '\n'
                           '“Can we build it again?”\n'
                           '\n'
                           'Kathy looked at the blue block.\n'
                           '\n'
                           'Then she looked at Kathin.\n'
                           '\n'
                           '“You can use the blue block for the door.”\n'
                           '\n'
                           'Kathin smiled.\n'
                           '\n'
                           '“And you can build the top?”\n'
                           '\n'
                           '“Okay.”\n'
                           '\n'
                           'They sat on the floor.\n'
                           '\n'
                           'One block.\n'
                           '\n'
                           'Two blocks.\n'
                           '\n'
                           'Three blocks.\n'
                           '\n'
                           'This time, they built something different.\n'
                           '\n'
                           'Not a tower.\n'
                           '\n'
                           'A lighthouse.\n'
                           '\n'
                           'Kathy made the top.\n'
                           '\n'
                           'Kathin made the door.\n'
                           '\n'
                           'And right in the middle, they placed two tiny blocks.\n'
                           '\n'
                           'One red.\n'
                           '\n'
                           'One blue.\n'
                           '\n'
                           'Mom walked past the room.\n'
                           '\n'
                           'She stopped at the door.\n'
                           '\n'
                           '“Nice lighthouse,” she said.\n'
                           '\n'
                           'Kathin smiled.\n'
                           '\n'
                           '“It has two lights.”\n'
                           '\n'
                           'Kathy added,\n'
                           '\n'
                           '“They still fight sometimes.”\n'
                           '\n'
                           'Mom smiled.\n'
                           '\n'
                           '“That sounds very realistic.”\n'
                           '\n'
                           'Kathy and Kathin laughed.\n'
                           '\n'
                           'And for the rest of the afternoon, the lighthouse stayed standing.\n'
                           '\n'
                           'Well...\n'
                           '\n'
                           'Almost.\n'
                           '\n'
                           'After dinner, Kathin bumped it with his foot.\n'
                           '\n'
                           'CRASH!\n'
                           '\n'
                           'Kathy looked at him.\n'
                           '\n'
                           'Kathin froze.\n'
                           '\n'
                           'Kathy took a big breath.\n'
                           '\n'
                           'Kathin took a big breath too.\n'
                           '\n'
                           'Then Kathin whispered,\n'
                           '\n'
                           '“Boat?”\n'
                           '\n'
                           'Kathy tried not to laugh.\n'
                           '\n'
                           '“Boat.”\n'
                           '\n'
                           'And together, they started building again.\n'
                           '\n'
                           'The End.'}],
  'audience': 'Kids'},
 {'book_id': 'the-velveteen-rabbit',
  'title': 'The Velveteen Rabbit',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🐰',
  'chapters': [{'chapter_id': 'the-velveteen-rabbit',
                'chapter_title': 'The Velveteen Rabbit',
                'order': 1,
                'content': 'Once upon a time, there was a little toy rabbit.\n'
                           '\n'
                           'He was made of soft brown velvet.\n'
                           '\n'
                           'He had long ears, a round body, and shiny eyes.\n'
                           '\n'
                           'On Christmas morning, a little boy found the rabbit in his Christmas '
                           'stocking.\n'
                           '\n'
                           '“Oh! A rabbit!” the boy cried.\n'
                           '\n'
                           'For a few hours, the boy loved him very much.\n'
                           '\n'
                           'He hugged him.\n'
                           '\n'
                           'He carried him around.\n'
                           '\n'
                           'He even took him to breakfast.\n'
                           '\n'
                           'But soon, the boy received many other presents.\n'
                           '\n'
                           'There were toy cars, trains, soldiers, and wonderful toys that could move '
                           'by themselves.\n'
                           '\n'
                           'The little rabbit could not move.\n'
                           '\n'
                           'He could not talk.\n'
                           '\n'
                           'He could not make sounds.\n'
                           '\n'
                           'So after a while, the boy forgot about him.\n'
                           '\n'
                           'The rabbit was placed in the toy cupboard.\n'
                           '\n'
                           'There he met many other toys.\n'
                           '\n'
                           'Some of the expensive toys were very proud.\n'
                           '\n'
                           '“I am made of metal,” said one.\n'
                           '\n'
                           '“I have real wheels,” said another.\n'
                           '\n'
                           'The little rabbit felt embarrassed.\n'
                           '\n'
                           'He knew he was only made of cloth and stuffing.\n'
                           '\n'
                           'But there was one old toy in the cupboard who was very kind.\n'
                           '\n'
                           'His name was the Skin Horse.\n'
                           '\n'
                           'The Skin Horse was old.\n'
                           '\n'
                           'His hair was thin.\n'
                           '\n'
                           'His body was worn.\n'
                           '\n'
                           'But he had been loved by children for many years.\n'
                           '\n'
                           'One day, the little rabbit asked him a question.\n'
                           '\n'
                           '“What does it mean to be real?”\n'
                           '\n'
                           'The Skin Horse thought for a moment.\n'
                           '\n'
                           '“Real is not about how you are made,” he said.\n'
                           '\n'
                           '“It happens when someone loves you for a very long time.”\n'
                           '\n'
                           'The rabbit looked surprised.\n'
                           '\n'
                           '“Does it hurt?”\n'
                           '\n'
                           '“Sometimes,” said the Skin Horse.\n'
                           '\n'
                           '“Your fur may become worn. Your eyes may become dull. You may not look new '
                           'anymore.”\n'
                           '\n'
                           'The rabbit became worried.\n'
                           '\n'
                           '“But when you are real, you do not mind looking old.”\n'
                           '\n'
                           '“Why?”\n'
                           '\n'
                           '“Because you know you are loved.”\n'
                           '\n'
                           'The little rabbit thought about this for a long time.\n'
                           '\n'
                           'He wanted very much to become real.\n'
                           '\n'
                           'One night, the boy could not find the toy he usually slept with.\n'
                           '\n'
                           'His mother looked around the room.\n'
                           '\n'
                           'Then she saw the little rabbit.\n'
                           '\n'
                           '“Here,” she said. “Take your old bunny.”\n'
                           '\n'
                           'The boy hugged the rabbit.\n'
                           '\n'
                           'That night, the rabbit slept beside him.\n'
                           '\n'
                           'The next night, the boy asked for him again.\n'
                           '\n'
                           'And the next night.\n'
                           '\n'
                           'Soon, the rabbit went everywhere with the boy.\n'
                           '\n'
                           'They played in the garden.\n'
                           '\n'
                           'They sat under trees.\n'
                           '\n'
                           'They made little houses in the grass.\n'
                           '\n'
                           'Sometimes the boy carried him by one ear.\n'
                           '\n'
                           'Sometimes he hugged him very tightly.\n'
                           '\n'
                           "The rabbit's soft velvet became worn.\n"
                           '\n'
                           'His beautiful fur began to disappear.\n'
                           '\n'
                           'His shape changed.\n'
                           '\n'
                           'But the rabbit did not care.\n'
                           '\n'
                           'He was happy.\n'
                           '\n'
                           'One day, the boy said,\n'
                           '\n'
                           '“My bunny is not a toy. He is real!”\n'
                           '\n'
                           "The rabbit's heart felt very warm.\n"
                           '\n'
                           'Perhaps the Skin Horse had been right.\n'
                           '\n'
                           'Perhaps love was making him real.\n'
                           '\n'
                           'In the garden, the rabbit once saw two real rabbits.\n'
                           '\n'
                           'They jumped around him.\n'
                           '\n'
                           '“Come and play!” they said.\n'
                           '\n'
                           "“I can't jump like you,” said the toy rabbit.\n"
                           '\n'
                           '“Why not?”\n'
                           '\n'
                           "“I don't have strong legs.”\n"
                           '\n'
                           'The wild rabbits laughed.\n'
                           '\n'
                           '“You are not a real rabbit!”\n'
                           '\n'
                           '“Yes, I am,” said the little rabbit.\n'
                           '\n'
                           '“The boy says I am real.”\n'
                           '\n'
                           'But the wild rabbits jumped away.\n'
                           '\n'
                           'The little rabbit felt sad.\n'
                           '\n'
                           'Still, when the boy came running to him, everything felt right again.\n'
                           '\n'
                           'The boy loved him.\n'
                           '\n'
                           'That was enough.\n'
                           '\n'
                           'Then, one day, the boy became very sick.\n'
                           '\n'
                           'He had a high fever.\n'
                           '\n'
                           'He stayed in bed for many days.\n'
                           '\n'
                           'The little rabbit stayed beside him the whole time.\n'
                           '\n'
                           'He wished he could make the boy feel better.\n'
                           '\n'
                           'At last, the boy became well again.\n'
                           '\n'
                           'Everyone was very happy.\n'
                           '\n'
                           "But the doctor told the boy's mother that everything the child had used "
                           'while he was sick had to be taken away.\n'
                           '\n'
                           'The sheets.\n'
                           '\n'
                           'The pillows.\n'
                           '\n'
                           'And the old toys.\n'
                           '\n'
                           'The little rabbit was placed in a bag with other things from the sickroom.\n'
                           '\n'
                           'He was carried outside to the garden.\n'
                           '\n'
                           'Tomorrow, everything in the bag would be burned.\n'
                           '\n'
                           'The rabbit lay alone under the night sky.\n'
                           '\n'
                           'He thought about the boy.\n'
                           '\n'
                           'He remembered their games.\n'
                           '\n'
                           'He remembered the warm bed.\n'
                           '\n'
                           'He remembered the boy saying,\n'
                           '\n'
                           '“You are real.”\n'
                           '\n'
                           "A tear fell from the rabbit's eye.\n"
                           '\n'
                           'The tear landed on the ground.\n'
                           '\n'
                           'And from that place, a beautiful flower appeared.\n'
                           '\n'
                           'Suddenly, a fairy came out of the flower.\n'
                           '\n'
                           'She looked at the little rabbit.\n'
                           '\n'
                           '“Why are you crying?” she asked.\n'
                           '\n'
                           '“I was loved,” said the rabbit.\n'
                           '\n'
                           '“And now I must leave my boy.”\n'
                           '\n'
                           'The fairy smiled.\n'
                           '\n'
                           '“You became real to the boy because he loved you.”\n'
                           '\n'
                           '“Yes,” said the rabbit.\n'
                           '\n'
                           '“But now I will make you real to everyone.”\n'
                           '\n'
                           'She kissed him.\n'
                           '\n'
                           'Suddenly, the rabbit felt something strange.\n'
                           '\n'
                           'His legs became strong.\n'
                           '\n'
                           'His body became warm.\n'
                           '\n'
                           'His ears moved.\n'
                           '\n'
                           'He jumped.\n'
                           '\n'
                           'For the first time in his life, he really jumped!\n'
                           '\n'
                           'The rabbit looked down.\n'
                           '\n'
                           'He was no longer a toy.\n'
                           '\n'
                           'He was a real rabbit.\n'
                           '\n'
                           'He ran into the forest with the other rabbits.\n'
                           '\n'
                           'Many months later, the boy was playing in the garden.\n'
                           '\n'
                           'He saw a brown rabbit watching him.\n'
                           '\n'
                           'The boy stopped.\n'
                           '\n'
                           '“That rabbit looks like my old bunny,” he said.\n'
                           '\n'
                           'The real rabbit looked at the boy.\n'
                           '\n'
                           'For one quiet moment, he remembered everything.\n'
                           '\n'
                           'The warm bed.\n'
                           '\n'
                           'The garden.\n'
                           '\n'
                           'The hugs.\n'
                           '\n'
                           'The love.\n'
                           '\n'
                           'Then he turned and jumped back into the forest.\n'
                           '\n'
                           'And the rabbit understood something important.\n'
                           '\n'
                           'Love had made him real.'}],
  'audience': 'Kids'},
 {'book_id': 'the-selfish-giant',
  'title': 'The Selfish Giant',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🌳',
  'chapters': [{'chapter_id': 'the-selfish-giant',
                'chapter_title': 'The Selfish Giant',
                'order': 1,
                'content': 'Once upon a time, there was a beautiful garden.\n'
                           '\n'
                           'The garden belonged to a giant.\n'
                           '\n'
                           'But the giant had been away for many years.\n'
                           '\n'
                           'So every afternoon, children came to play there.\n'
                           '\n'
                           'The garden was wonderful.\n'
                           '\n'
                           'There was soft green grass.\n'
                           '\n'
                           'There were beautiful flowers.\n'
                           '\n'
                           'There were peach trees that became pink and white in spring.\n'
                           '\n'
                           'Birds sat in the trees and sang.\n'
                           '\n'
                           'The children loved the garden.\n'
                           '\n'
                           '“This is the happiest place in the world,” they said.\n'
                           '\n'
                           'One day, the giant came home.\n'
                           '\n'
                           'He saw the children playing.\n'
                           '\n'
                           '“What are you doing in my garden?” he shouted.\n'
                           '\n'
                           'The children became frightened and ran away.\n'
                           '\n'
                           '“This garden belongs to me!”\n'
                           '\n'
                           'The giant built a tall wall around it.\n'
                           '\n'
                           'Then he put up a big sign.\n'
                           '\n'
                           'NO CHILDREN ALLOWED.\n'
                           '\n'
                           'The children were very sad.\n'
                           '\n'
                           'They had nowhere beautiful to play.\n'
                           '\n'
                           'Then spring came.\n'
                           '\n'
                           'Flowers appeared everywhere.\n'
                           '\n'
                           'Birds sang.\n'
                           '\n'
                           'Trees became green.\n'
                           '\n'
                           "But inside the giant's garden, it was still winter.\n"
                           '\n'
                           'Snow covered the grass.\n'
                           '\n'
                           'The trees had no leaves.\n'
                           '\n'
                           'The birds did not sing.\n'
                           '\n'
                           '“Why is spring so late?” wondered the giant.\n'
                           '\n'
                           'Summer came.\n'
                           '\n'
                           "But not to the giant's garden.\n"
                           '\n'
                           'Autumn came.\n'
                           '\n'
                           'But the trees gave the giant no fruit.\n'
                           '\n'
                           '“Why does my garden stay cold?” he asked.\n'
                           '\n'
                           'One morning, the giant heard a beautiful sound.\n'
                           '\n'
                           'A little bird was singing outside his window.\n'
                           '\n'
                           'The giant jumped out of bed.\n'
                           '\n'
                           '“Has spring finally come?”\n'
                           '\n'
                           'He looked outside.\n'
                           '\n'
                           'Something wonderful had happened.\n'
                           '\n'
                           'The children had found a small hole in the wall.\n'
                           '\n'
                           'They had climbed into the garden.\n'
                           '\n'
                           'There was a child sitting in every tree.\n'
                           '\n'
                           'And wherever a child sat, the tree began to bloom.\n'
                           '\n'
                           'Flowers opened.\n'
                           '\n'
                           'Birds returned.\n'
                           '\n'
                           'The grass became green.\n'
                           '\n'
                           'Spring had come back.\n'
                           '\n'
                           'But in one corner of the garden, it was still winter.\n'
                           '\n'
                           'A very small boy stood under a tree.\n'
                           '\n'
                           'He was too little to climb.\n'
                           '\n'
                           'The tree bent its branches down, but the boy still could not reach them.\n'
                           '\n'
                           'The boy began to cry.\n'
                           '\n'
                           'The giant watched from his window.\n'
                           '\n'
                           'Suddenly, his heart became soft.\n'
                           '\n'
                           '“How selfish I have been,” he said.\n'
                           '\n'
                           '“Now I understand why spring would not come.”\n'
                           '\n'
                           'The giant went outside.\n'
                           '\n'
                           'When the children saw him, they were frightened.\n'
                           '\n'
                           'They ran away.\n'
                           '\n'
                           'And winter returned.\n'
                           '\n'
                           'But the little boy did not see the giant.\n'
                           '\n'
                           'His eyes were full of tears.\n'
                           '\n'
                           'The giant walked gently toward him.\n'
                           '\n'
                           'He picked him up.\n'
                           '\n'
                           'Then he placed him in the tree.\n'
                           '\n'
                           'At once, the tree filled with flowers.\n'
                           '\n'
                           'Birds began to sing.\n'
                           '\n'
                           'The little boy smiled.\n'
                           '\n'
                           "Then he put his arms around the giant's neck and kissed him.\n"
                           '\n'
                           'The other children saw that the giant was no longer angry.\n'
                           '\n'
                           'They came running back.\n'
                           '\n'
                           'The giant smiled.\n'
                           '\n'
                           '“This is your garden now too,” he said.\n'
                           '\n'
                           'Then he took a large hammer.\n'
                           '\n'
                           'He knocked down the wall.\n'
                           '\n'
                           'From that day on, the children played in the garden every afternoon.\n'
                           '\n'
                           'And the giant played with them.\n'
                           '\n'
                           'He was happy.\n'
                           '\n'
                           'But the giant often looked for the little boy who had kissed him.\n'
                           '\n'
                           '“Where is your little friend?” he asked the other children.\n'
                           '\n'
                           "“We don't know,” they said.\n"
                           '\n'
                           'Years passed.\n'
                           '\n'
                           'The giant became old.\n'
                           '\n'
                           'He could no longer run and play.\n'
                           '\n'
                           'So he sat in a large chair and watched the children.\n'
                           '\n'
                           '“I have many beautiful flowers,” he said.\n'
                           '\n'
                           '“But the children are the most beautiful flowers of all.”\n'
                           '\n'
                           'One winter morning, the giant looked outside.\n'
                           '\n'
                           'He could not believe his eyes.\n'
                           '\n'
                           'In one corner of the garden, a tree was covered with beautiful white '
                           'flowers.\n'
                           '\n'
                           'Under the tree stood the little boy.\n'
                           '\n'
                           'The giant was filled with joy.\n'
                           '\n'
                           'He hurried outside.\n'
                           '\n'
                           '“You came back!” he said.\n'
                           '\n'
                           'The little boy smiled.\n'
                           '\n'
                           '“You let me play in your garden long ago.”\n'
                           '\n'
                           '“Yes,” said the giant.\n'
                           '\n'
                           '“Today, you will come and play in my garden.”\n'
                           '\n'
                           'The giant felt peaceful.\n'
                           '\n'
                           'Later that afternoon, the children came to play.\n'
                           '\n'
                           'They found the old giant resting quietly under the tree.\n'
                           '\n'
                           'White flowers covered him gently.\n'
                           '\n'
                           'The children remembered him with love.\n'
                           '\n'
                           'And the garden remained open forever.'}],
  'audience': 'Kids'},
 {'book_id': 'the-happy-prince',
  'title': 'The Happy Prince',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🕊️',
  'chapters': [{'chapter_id': 'the-happy-prince',
                'chapter_title': 'The Happy Prince',
                'order': 1,
                'content': 'High above a city stood a beautiful statue.\n'
                           '\n'
                           'It was called the Happy Prince.\n'
                           '\n'
                           'His body was covered with gold.\n'
                           '\n'
                           'He had two bright jewels for eyes.\n'
                           '\n'
                           'A large red jewel shone on his sword.\n'
                           '\n'
                           'Everyone admired him.\n'
                           '\n'
                           '“He looks so happy,” people said.\n'
                           '\n'
                           'One evening, a little swallow flew over the city.\n'
                           '\n'
                           'The other swallows had already flown south for the winter.\n'
                           '\n'
                           'But this swallow had stayed behind.\n'
                           '\n'
                           'Now he was finally going to join them.\n'
                           '\n'
                           'He decided to sleep at the feet of the Happy Prince.\n'
                           '\n'
                           '“This is a wonderful place to sleep,” he said.\n'
                           '\n'
                           'But suddenly—\n'
                           '\n'
                           'SPLASH!\n'
                           '\n'
                           'A drop of water fell on him.\n'
                           '\n'
                           '“Rain?”\n'
                           '\n'
                           'The swallow looked up.\n'
                           '\n'
                           'There were no clouds.\n'
                           '\n'
                           'Another drop fell.\n'
                           '\n'
                           'Then he saw that the Happy Prince was crying.\n'
                           '\n'
                           '“Why are you crying?” asked the swallow.\n'
                           '\n'
                           '“When I was alive,” said the Prince, “I lived inside a palace.”\n'
                           '\n'
                           '“I never saw sadness.”\n'
                           '\n'
                           '“I thought everyone was happy.”\n'
                           '\n'
                           '“But now I stand high above the city, and I can see everything.”\n'
                           '\n'
                           'The Prince looked toward a small house.\n'
                           '\n'
                           '“There is a poor mother inside.”\n'
                           '\n'
                           '“She works very hard.”\n'
                           '\n'
                           '“Her little boy is sick.”\n'
                           '\n'
                           '“She has no money to buy what he needs.”\n'
                           '\n'
                           '“Little swallow, please take the red jewel from my sword and give it to '
                           'her.”\n'
                           '\n'
                           'The swallow hesitated.\n'
                           '\n'
                           'He wanted to fly south.\n'
                           '\n'
                           'His friends were waiting.\n'
                           '\n'
                           'But the Prince looked so sad.\n'
                           '\n'
                           '“All right,” said the swallow.\n'
                           '\n'
                           '“Just for one night.”\n'
                           '\n'
                           'He took the red jewel.\n'
                           '\n'
                           'He flew across the city.\n'
                           '\n'
                           'He found the little house.\n'
                           '\n'
                           'He placed the jewel near the mother.\n'
                           '\n'
                           'Then he flew back.\n'
                           '\n'
                           '“It is strange,” said the swallow.\n'
                           '\n'
                           '“The night is cold, but I feel warm.”\n'
                           '\n'
                           '“That is because you did something kind,” said the Prince.\n'
                           '\n'
                           'The next night, the Prince asked for help again.\n'
                           '\n'
                           'He could see a young man sitting in a cold room.\n'
                           '\n'
                           'The man was trying to write.\n'
                           '\n'
                           'But he had no fire.\n'
                           '\n'
                           'He was hungry.\n'
                           '\n'
                           '“Take one of my eyes,” said the Prince.\n'
                           '\n'
                           '“No!” cried the swallow.\n'
                           '\n'
                           '“Please.”\n'
                           '\n'
                           "The swallow sadly removed one of the Prince's jewel eyes.\n"
                           '\n'
                           'He took it to the young man.\n'
                           '\n'
                           'The next day, the Prince saw a little girl selling matches.\n'
                           '\n'
                           'Her matches had fallen into the water.\n'
                           '\n'
                           'She was afraid to go home because she had earned no money.\n'
                           '\n'
                           '“Take my other eye,” said the Prince.\n'
                           '\n'
                           '“But then you will be blind!”\n'
                           '\n'
                           '“Please,” said the Prince.\n'
                           '\n'
                           "So the swallow took the Prince's second eye and gave it to the girl.\n"
                           '\n'
                           'Now the Happy Prince could not see.\n'
                           '\n'
                           '“I will stay with you,” said the swallow.\n'
                           '\n'
                           'The weather became colder.\n'
                           '\n'
                           'The Prince asked the swallow to fly around the city.\n'
                           '\n'
                           'The swallow told him what he saw.\n'
                           '\n'
                           'Poor children.\n'
                           '\n'
                           'Hungry families.\n'
                           '\n'
                           'People sleeping in cold streets.\n'
                           '\n'
                           'The Prince said,\n'
                           '\n'
                           '“Take the gold from my body.”\n'
                           '\n'
                           'So the swallow removed the gold, piece by piece.\n'
                           '\n'
                           'He gave it to the poor.\n'
                           '\n'
                           'Soon, children had food.\n'
                           '\n'
                           'Families had warmth.\n'
                           '\n'
                           'People smiled again.\n'
                           '\n'
                           'But the beautiful Happy Prince was no longer beautiful.\n'
                           '\n'
                           'He was gray.\n'
                           '\n'
                           'His jewels were gone.\n'
                           '\n'
                           'His gold was gone.\n'
                           '\n'
                           'Winter became colder.\n'
                           '\n'
                           'The swallow knew he should leave.\n'
                           '\n'
                           'But he loved the Prince.\n'
                           '\n'
                           'At last, the swallow became too cold to fly.\n'
                           '\n'
                           "He sat at the Prince's feet.\n"
                           '\n'
                           '“Goodbye, dear Prince,” he whispered.\n'
                           '\n'
                           'Then the little swallow died.\n'
                           '\n'
                           'At that moment, something inside the statue broke.\n'
                           '\n'
                           "It was the Prince's lead heart.\n"
                           '\n'
                           'The next morning, the people looked at the statue.\n'
                           '\n'
                           '“Oh!” they said.\n'
                           '\n'
                           '“The Happy Prince is ugly now!”\n'
                           '\n'
                           'They took the statue down.\n'
                           '\n'
                           'They melted the metal.\n'
                           '\n'
                           'But the broken lead heart would not melt.\n'
                           '\n'
                           'So they threw it away.\n'
                           '\n'
                           'They threw the little swallow away too.\n'
                           '\n'
                           'But in the story, those two things were the most precious things in the '
                           'whole city.\n'
                           '\n'
                           'Because one had given everything he had.\n'
                           '\n'
                           'And the other had stayed because of love.'}],
  'audience': 'Kids'},
 {'book_id': 'the-ugly-duckling',
  'title': 'The Ugly Duckling',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🦢',
  'chapters': [{'chapter_id': 'the-ugly-duckling',
                'chapter_title': 'The Ugly Duckling',
                'order': 1,
                'content': 'Once upon a time, a mother duck sat on her eggs.\n'
                           '\n'
                           'She waited patiently.\n'
                           '\n'
                           'One morning—\n'
                           '\n'
                           'Crack!\n'
                           '\n'
                           'Crack!\n'
                           '\n'
                           'Crack!\n'
                           '\n'
                           'Little ducklings came out.\n'
                           '\n'
                           '“Peep! Peep!”\n'
                           '\n'
                           'They were small and yellow.\n'
                           '\n'
                           'But one egg was much bigger than the others.\n'
                           '\n'
                           'At last—\n'
                           '\n'
                           'CRACK!\n'
                           '\n'
                           'A large gray duckling came out.\n'
                           '\n'
                           'The mother duck looked at him.\n'
                           '\n'
                           '“He is very big,” she said.\n'
                           '\n'
                           'The other ducklings stared.\n'
                           '\n'
                           '“He looks strange.”\n'
                           '\n'
                           'Soon, they began teasing him.\n'
                           '\n'
                           '“You are ugly!”\n'
                           '\n'
                           "“You don't look like us!”\n"
                           '\n'
                           'The gray duckling became very sad.\n'
                           '\n'
                           'Even some of the other animals were unkind.\n'
                           '\n'
                           'The duckling wondered,\n'
                           '\n'
                           '“Why am I different?”\n'
                           '\n'
                           'One day, he could not take it anymore.\n'
                           '\n'
                           'He ran away.\n'
                           '\n'
                           'He traveled far from home.\n'
                           '\n'
                           'He found wild ducks.\n'
                           '\n'
                           'But they laughed at him too.\n'
                           '\n'
                           'He found a little house where an old woman lived with a cat and a hen.\n'
                           '\n'
                           '“You can stay here,” said the old woman.\n'
                           '\n'
                           'But the cat and the hen did not understand him.\n'
                           '\n'
                           'The duckling loved swimming.\n'
                           '\n'
                           '“I love the water,” he said.\n'
                           '\n'
                           'The hen laughed.\n'
                           '\n'
                           '“Swimming? What a silly thing!”\n'
                           '\n'
                           'The duckling felt lonely again.\n'
                           '\n'
                           'So he left.\n'
                           '\n'
                           'Autumn came.\n'
                           '\n'
                           'The leaves fell from the trees.\n'
                           '\n'
                           'One evening, the duckling looked into the sky.\n'
                           '\n'
                           'He saw a group of beautiful white birds.\n'
                           '\n'
                           'They had long necks.\n'
                           '\n'
                           'They flew gracefully through the sky.\n'
                           '\n'
                           'The duckling had never seen anything so beautiful.\n'
                           '\n'
                           '“I wish I could be like them,” he thought.\n'
                           '\n'
                           'Winter came.\n'
                           '\n'
                           'It was very cold.\n'
                           '\n'
                           'The poor duckling struggled to survive.\n'
                           '\n'
                           'But eventually spring arrived.\n'
                           '\n'
                           'The sun became warm again.\n'
                           '\n'
                           'Flowers began to grow.\n'
                           '\n'
                           'The duckling saw the beautiful white birds again.\n'
                           '\n'
                           'Swans.\n'
                           '\n'
                           'He wanted to go near them.\n'
                           '\n'
                           'But he was afraid.\n'
                           '\n'
                           '“They will laugh at me too,” he thought.\n'
                           '\n'
                           'Still, he swam toward them.\n'
                           '\n'
                           '“If they do not like me, I will leave.”\n'
                           '\n'
                           'The duckling lowered his head.\n'
                           '\n'
                           'Then he saw his reflection in the water.\n'
                           '\n'
                           'He froze.\n'
                           '\n'
                           'The gray feathers were gone.\n'
                           '\n'
                           'His neck was long and graceful.\n'
                           '\n'
                           'His feathers were white.\n'
                           '\n'
                           'He was not an ugly duckling.\n'
                           '\n'
                           'He was a beautiful swan.\n'
                           '\n'
                           'The other swans swam toward him.\n'
                           '\n'
                           '“Welcome,” they said.\n'
                           '\n'
                           'Children near the lake pointed at him.\n'
                           '\n'
                           '“Look! A new swan!”\n'
                           '\n'
                           '“He is beautiful!”\n'
                           '\n'
                           'The young swan remembered all the times people had laughed at him.\n'
                           '\n'
                           'But now he understood.\n'
                           '\n'
                           'There had never been anything wrong with him.\n'
                           '\n'
                           'He had simply been different.\n'
                           '\n'
                           'And he had needed time to grow into himself.\n'
                           '\n'
                           'He lifted his wings.\n'
                           '\n'
                           'For the first time, he felt happy to be exactly who he was.'}],
  'audience': 'Kids'},
 {'book_id': 'the-emperors-new-clothes',
  'title': 'The Emperor’s New Clothes',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '👑',
  'chapters': [{'chapter_id': 'the-emperors-new-clothes',
                'chapter_title': 'The Emperor’s New Clothes',
                'order': 1,
                'content': 'Once upon a time, there was an emperor who loved clothes.\n'
                           '\n'
                           'He did not care much about his soldiers.\n'
                           '\n'
                           'He did not care much about meetings.\n'
                           '\n'
                           'He cared about clothes.\n'
                           '\n'
                           'Every hour, he changed into something new.\n'
                           '\n'
                           'One day, two dishonest men came to the city.\n'
                           '\n'
                           '“We make the most amazing clothes in the world,” they said.\n'
                           '\n'
                           'The emperor became excited.\n'
                           '\n'
                           '“What is special about them?”\n'
                           '\n'
                           'The men smiled.\n'
                           '\n'
                           '“Our cloth is magical.”\n'
                           '\n'
                           '“Magical?”\n'
                           '\n'
                           '“Yes. Foolish people cannot see it.”\n'
                           '\n'
                           'The emperor thought,\n'
                           '\n'
                           '“If I wear these clothes, I will know who is clever and who is foolish!”\n'
                           '\n'
                           'He paid the men a lot of money.\n'
                           '\n'
                           'The two men put empty frames in a room.\n'
                           '\n'
                           'They pretended to weave.\n'
                           '\n'
                           'But there was nothing there.\n'
                           '\n'
                           'After a few days, the emperor sent one of his officials to check.\n'
                           '\n'
                           'The official looked at the empty frame.\n'
                           '\n'
                           'He saw nothing.\n'
                           '\n'
                           '“Oh no!” he thought.\n'
                           '\n'
                           '“Does this mean I am foolish?”\n'
                           '\n'
                           'The two men asked,\n'
                           '\n'
                           "“Isn't the cloth beautiful?”\n"
                           '\n'
                           'The official was afraid to tell the truth.\n'
                           '\n'
                           '“Yes!” he said.\n'
                           '\n'
                           '“It is wonderful!”\n'
                           '\n'
                           'He went back to the emperor.\n'
                           '\n'
                           '“The cloth is beautiful,” he said.\n'
                           '\n'
                           'Later, another official visited.\n'
                           '\n'
                           'He also saw nothing.\n'
                           '\n'
                           'But he was afraid.\n'
                           '\n'
                           '“Wonderful!” he said.\n'
                           '\n'
                           'Finally, the emperor went to see the cloth himself.\n'
                           '\n'
                           'He looked.\n'
                           '\n'
                           'There was nothing.\n'
                           '\n'
                           'His heart jumped.\n'
                           '\n'
                           "“I can't see anything!”\n"
                           '\n'
                           'Then he thought,\n'
                           '\n'
                           '“Perhaps I am foolish.”\n'
                           '\n'
                           'So he smiled.\n'
                           '\n'
                           '“Beautiful!” he cried.\n'
                           '\n'
                           '“Make me a suit for the big parade!”\n'
                           '\n'
                           'The two men pretended to cut the invisible cloth.\n'
                           '\n'
                           'They pretended to sew it.\n'
                           '\n'
                           'On the morning of the parade, they said,\n'
                           '\n'
                           '“Your Majesty, your new clothes are ready.”\n'
                           '\n'
                           'The emperor took off his old clothes.\n'
                           '\n'
                           'The men pretended to dress him.\n'
                           '\n'
                           '“How light they are!” they said.\n'
                           '\n'
                           'The emperor looked in the mirror.\n'
                           '\n'
                           'He saw himself wearing nothing.\n'
                           '\n'
                           'But he did not want anyone to think he was foolish.\n'
                           '\n'
                           '“Perfect!” he said.\n'
                           '\n'
                           'The parade began.\n'
                           '\n'
                           'The emperor walked through the streets.\n'
                           '\n'
                           'Everyone had heard about the magical clothes.\n'
                           '\n'
                           'So although nobody could see them, everyone shouted,\n'
                           '\n'
                           '“Beautiful!”\n'
                           '\n'
                           '“Wonderful!”\n'
                           '\n'
                           '“What amazing clothes!”\n'
                           '\n'
                           'Nobody wanted to be called foolish.\n'
                           '\n'
                           'Then a little child looked at the emperor.\n'
                           '\n'
                           'The child did not know about pretending.\n'
                           '\n'
                           'The child did not care about looking clever.\n'
                           '\n'
                           'So the child simply said,\n'
                           '\n'
                           "“But he isn't wearing any clothes!”\n"
                           '\n'
                           'The street became quiet.\n'
                           '\n'
                           'Then someone whispered,\n'
                           '\n'
                           '“The child is right.”\n'
                           '\n'
                           'Another person said,\n'
                           '\n'
                           "“He really isn't wearing anything.”\n"
                           '\n'
                           'Soon everyone was talking.\n'
                           '\n'
                           'The emperor knew they were right.\n'
                           '\n'
                           'He felt embarrassed.\n'
                           '\n'
                           'But he continued walking.\n'
                           '\n'
                           'And perhaps, for the first time, he learned something important.\n'
                           '\n'
                           'Sometimes many people can pretend something is true.\n'
                           '\n'
                           'It takes courage to say what you really see.'}],
  'audience': 'Kids'},
 {'book_id': 'the-lion-and-the-mouse',
  'title': 'The Lion and the Mouse',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🦁',
  'chapters': [{'chapter_id': 'the-lion-and-the-mouse',
                'chapter_title': 'The Lion and the Mouse',
                'order': 1,
                'content': 'One warm afternoon, a lion was sleeping under a tree.\n'
                           '\n'
                           'He was the strongest animal in the forest.\n'
                           '\n'
                           'A tiny mouse ran across the ground.\n'
                           '\n'
                           'The mouse did not see the lion.\n'
                           '\n'
                           "Suddenly, he ran right across the lion's paw.\n"
                           '\n'
                           'The lion woke up.\n'
                           '\n'
                           'ROAR!\n'
                           '\n'
                           'He caught the mouse.\n'
                           '\n'
                           '“How dare you wake me!”\n'
                           '\n'
                           'The mouse trembled.\n'
                           '\n'
                           "“Please don't eat me!”\n"
                           '\n'
                           'The lion laughed.\n'
                           '\n'
                           '“You are so small!”\n'
                           '\n'
                           '“Please let me go,” said the mouse.\n'
                           '\n'
                           '“Maybe one day I can help you.”\n'
                           '\n'
                           'The lion laughed even louder.\n'
                           '\n'
                           '“You? Help me?”\n'
                           '\n'
                           'But the lion was in a good mood.\n'
                           '\n'
                           'So he opened his paw.\n'
                           '\n'
                           '“Go.”\n'
                           '\n'
                           'The mouse ran away.\n'
                           '\n'
                           '“Thank you!” he called.\n'
                           '\n'
                           'A few days later, hunters came to the forest.\n'
                           '\n'
                           'They placed a strong net between the trees.\n'
                           '\n'
                           'The lion walked into it.\n'
                           '\n'
                           'The net closed around him.\n'
                           '\n'
                           'He pulled.\n'
                           '\n'
                           'He pushed.\n'
                           '\n'
                           'He roared.\n'
                           '\n'
                           'But the more he moved, the tighter the net became.\n'
                           '\n'
                           '“Help!” roared the lion.\n'
                           '\n'
                           'Far away, the little mouse heard him.\n'
                           '\n'
                           '“I know that voice!”\n'
                           '\n'
                           'He ran toward the sound.\n'
                           '\n'
                           'He found the lion trapped in the net.\n'
                           '\n'
                           "“Don't worry,” said the mouse.\n"
                           '\n'
                           'The lion looked at him.\n'
                           '\n'
                           '“You?”\n'
                           '\n'
                           'The mouse began biting the rope.\n'
                           '\n'
                           'Nibble.\n'
                           '\n'
                           'Nibble.\n'
                           '\n'
                           'Nibble.\n'
                           '\n'
                           'The rope was thick.\n'
                           '\n'
                           'But the mouse did not stop.\n'
                           '\n'
                           'Nibble.\n'
                           '\n'
                           'Nibble.\n'
                           '\n'
                           'SNAP!\n'
                           '\n'
                           'One rope broke.\n'
                           '\n'
                           'Then another.\n'
                           '\n'
                           'And another.\n'
                           '\n'
                           'Finally, the lion was free.\n'
                           '\n'
                           'The lion looked at the tiny mouse.\n'
                           '\n'
                           '“I laughed when you said you could help me.”\n'
                           '\n'
                           'The mouse smiled.\n'
                           '\n'
                           '“Sometimes small friends can do big things.”\n'
                           '\n'
                           'The lion lowered his great head.\n'
                           '\n'
                           '“Thank you, my friend.”\n'
                           '\n'
                           'From that day on, the lion never judged someone by their size again.'}],
  'audience': 'Kids'},
 {'book_id': 'the-boy-who-cried-wolf',
  'title': 'The Boy Who Cried Wolf',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '🐺',
  'chapters': [{'chapter_id': 'the-boy-who-cried-wolf',
                'chapter_title': 'The Boy Who Cried Wolf',
                'order': 1,
                'content': 'Once upon a time, there was a young shepherd boy.\n'
                           '\n'
                           'Every day, he took the village sheep to a hill.\n'
                           '\n'
                           'His job was simple.\n'
                           '\n'
                           'Watch the sheep.\n'
                           '\n'
                           'Keep them safe.\n'
                           '\n'
                           'Call for help if a wolf came.\n'
                           '\n'
                           'At first, the boy enjoyed the job.\n'
                           '\n'
                           'But after many days, he became bored.\n'
                           '\n'
                           'The sheep ate grass.\n'
                           '\n'
                           'The clouds moved slowly.\n'
                           '\n'
                           'Nothing exciting happened.\n'
                           '\n'
                           'Then the boy had an idea.\n'
                           '\n'
                           'He ran toward the village.\n'
                           '\n'
                           '“Wolf! Wolf!”\n'
                           '\n'
                           '“A wolf is attacking the sheep!”\n'
                           '\n'
                           'The villagers dropped their work.\n'
                           '\n'
                           'They ran up the hill.\n'
                           '\n'
                           'Some carried sticks.\n'
                           '\n'
                           'Others carried tools.\n'
                           '\n'
                           'But when they reached the sheep, there was no wolf.\n'
                           '\n'
                           'The boy laughed.\n'
                           '\n'
                           '“There is no wolf!”\n'
                           '\n'
                           'The villagers were angry.\n'
                           '\n'
                           '“This is not funny,” they said.\n'
                           '\n'
                           'A few days later, the boy became bored again.\n'
                           '\n'
                           'He remembered how everyone had run to help him.\n'
                           '\n'
                           'He smiled.\n'
                           '\n'
                           '“Wolf! Wolf!”\n'
                           '\n'
                           'Once again, the villagers ran up the hill.\n'
                           '\n'
                           'Once again, there was no wolf.\n'
                           '\n'
                           'The boy laughed.\n'
                           '\n'
                           'The villagers were even angrier.\n'
                           '\n'
                           '“Do not lie about danger,” they warned him.\n'
                           '\n'
                           'The next week, the boy was watching the sheep.\n'
                           '\n'
                           'Suddenly, the sheep became nervous.\n'
                           '\n'
                           'The boy looked toward the trees.\n'
                           '\n'
                           'Two bright eyes were watching him.\n'
                           '\n'
                           'A wolf stepped out.\n'
                           '\n'
                           'This time, it was real.\n'
                           '\n'
                           "The boy's heart began to pound.\n"
                           '\n'
                           '“WOLF!”\n'
                           '\n'
                           'He ran toward the village.\n'
                           '\n'
                           '“Wolf! Please help!”\n'
                           '\n'
                           'But the villagers heard him and shook their heads.\n'
                           '\n'
                           '“He is playing another trick.”\n'
                           '\n'
                           '“No one should listen.”\n'
                           '\n'
                           'The boy shouted again.\n'
                           '\n'
                           '“Please! This time it is true!”\n'
                           '\n'
                           'Nobody came.\n'
                           '\n'
                           'The wolf frightened the sheep and scattered them across the hill.\n'
                           '\n'
                           'The boy spent hours searching for them.\n'
                           '\n'
                           'That evening, the boy returned home exhausted and sad.\n'
                           '\n'
                           'An old villager spoke gently to him.\n'
                           '\n'
                           '“Do you understand what happened?”\n'
                           '\n'
                           'The boy nodded.\n'
                           '\n'
                           '“When I told the truth, nobody believed me.”\n'
                           '\n'
                           '“Why?”\n'
                           '\n'
                           '“Because I lied before.”\n'
                           '\n'
                           'The old man nodded.\n'
                           '\n'
                           '“Trust takes time to build.”\n'
                           '\n'
                           'The boy never forgot that day.\n'
                           '\n'
                           'After that, when he spoke, he tried hard to tell the truth.\n'
                           '\n'
                           'And slowly, people began to trust him again.'}],
  'audience': 'Kids'},
 {'book_id': 'the-little-match-girl',
  'title': 'The Little Match Girl',
  'subtitle': "Children's Story",
  'author': '',
  'content_type': 'Book',
  'category': "Children's Story",
  'description': '',
  'cover_emoji': '✨',
  'chapters': [{'chapter_id': 'the-little-match-girl',
                'chapter_title': 'The Little Match Girl',
                'order': 1,
                'content': 'Once upon a time, on a very cold winter night, a little girl walked through '
                           'the streets.\n'
                           '\n'
                           'Snow fell all around her.\n'
                           '\n'
                           'It was the last night of the year.\n'
                           '\n'
                           'People were inside their warm homes.\n'
                           '\n'
                           'They ate delicious food.\n'
                           '\n'
                           'Candles shone through the windows.\n'
                           '\n'
                           'But the little girl was outside.\n'
                           '\n'
                           'She was poor.\n'
                           '\n'
                           'She carried a small box of matches.\n'
                           '\n'
                           '“Matches!” she called.\n'
                           '\n'
                           '“Please buy some matches!”\n'
                           '\n'
                           'But everyone walked past her.\n'
                           '\n'
                           "The little girl's hands were freezing.\n"
                           '\n'
                           'Her feet were cold.\n'
                           '\n'
                           'She had sold nothing all day.\n'
                           '\n'
                           'She was afraid to go home because she had no money.\n'
                           '\n'
                           'So she sat down between two buildings.\n'
                           '\n'
                           'She looked at the matches in her hand.\n'
                           '\n'
                           '“Maybe I can light just one,” she thought.\n'
                           '\n'
                           'She took out a match.\n'
                           '\n'
                           'SCRATCH!\n'
                           '\n'
                           'A warm flame appeared.\n'
                           '\n'
                           'Suddenly, the little girl imagined a beautiful fireplace.\n'
                           '\n'
                           'The fire was warm.\n'
                           '\n'
                           'She held out her hands.\n'
                           '\n'
                           'But then—\n'
                           '\n'
                           'The match went out.\n'
                           '\n'
                           'The fireplace disappeared.\n'
                           '\n'
                           'The girl lit another match.\n'
                           '\n'
                           'SCRATCH!\n'
                           '\n'
                           'This time, she saw a beautiful table.\n'
                           '\n'
                           'There was warm food.\n'
                           '\n'
                           'Bread.\n'
                           '\n'
                           'Fruit.\n'
                           '\n'
                           'A wonderful dinner.\n'
                           '\n'
                           'The little girl smiled.\n'
                           '\n'
                           'Then the match went out.\n'
                           '\n'
                           'The food disappeared.\n'
                           '\n'
                           'She lit another.\n'
                           '\n'
                           'SCRATCH!\n'
                           '\n'
                           'Now she saw a beautiful Christmas tree.\n'
                           '\n'
                           'It was covered with hundreds of lights.\n'
                           '\n'
                           'The lights seemed to rise into the sky.\n'
                           '\n'
                           'Then one light fell.\n'
                           '\n'
                           '“A star is falling,” she whispered.\n'
                           '\n'
                           'Her grandmother had once told her that when a star falls, a soul is going '
                           'to heaven.\n'
                           '\n'
                           'The little girl thought of her grandmother.\n'
                           '\n'
                           'Her grandmother had been the only person who had made her feel completely '
                           'safe and loved.\n'
                           '\n'
                           'The girl lit another match.\n'
                           '\n'
                           'SCRATCH!\n'
                           '\n'
                           'There stood her grandmother.\n'
                           '\n'
                           'Warm.\n'
                           '\n'
                           'Kind.\n'
                           '\n'
                           'Smiling.\n'
                           '\n'
                           '“Grandmother!”\n'
                           '\n'
                           'The little girl was so happy.\n'
                           '\n'
                           'She knew the match would soon go out.\n'
                           '\n'
                           "“Please don't leave me!”\n"
                           '\n'
                           'She quickly lit all the remaining matches.\n'
                           '\n'
                           'The street became filled with light.\n'
                           '\n'
                           'In that bright light, the girl imagined her grandmother holding her.\n'
                           '\n'
                           '“You are safe,” her grandmother seemed to say.\n'
                           '\n'
                           '“You are loved.”\n'
                           '\n'
                           'The little girl smiled.\n'
                           '\n'
                           'The next morning, people found her sitting quietly in the snow.\n'
                           '\n'
                           'The matches were gone.\n'
                           '\n'
                           'There was a peaceful smile on her face.\n'
                           '\n'
                           'People said,\n'
                           '\n'
                           '“She was trying to keep warm.”\n'
                           '\n'
                           'But they did not know about the wonderful things she had imagined.\n'
                           '\n'
                           'They did not see the warm fireplace.\n'
                           '\n'
                           'They did not see the beautiful dinner.\n'
                           '\n'
                           'They did not see the Christmas tree.\n'
                           '\n'
                           'And they did not see her grandmother.\n'
                           '\n'
                           'The story is a sad one.\n'
                           '\n'
                           'But it asks us to notice people who are cold, hungry, lonely, or '
                           'forgotten.\n'
                           '\n'
                           'Because sometimes the person who needs our kindness most is the person '
                           'everyone else walks past.\n'
                           '\n'
                           'The End.'}],
  'audience': 'Kids'}]


# =========================================================
# PERSISTENT STORAGE (SUPABASE)
# =========================================================
@st.cache_resource
def get_supabase_client():
    """Create one server-side Supabase client per Streamlit process."""
    try:
        url = str(st.secrets["SUPABASE_URL"]).strip()
        key = str(st.secrets["SUPABASE_KEY"]).strip()
    except Exception:
        return None
    if not url or not key:
        return None
    return create_client(url, key)


def storage_owner() -> str:
    try:
        owner = str(st.secrets["BUNNY_OWNER"]).strip()
    except Exception:
        owner = "polly"
    return owner or "polly"


def storage_is_configured() -> bool:
    return get_supabase_client() is not None


def load_persistent_state():
    """Load user-imported books, reading progress, and bookmarks.

    Built-in DEMO_BOOKS remain in source code; only user-created data and
    user state are stored in Supabase.
    """
    client = get_supabase_client()
    if client is None:
        return {"books": [], "last_read": {}, "bookmarks": []}, None

    owner = storage_owner()
    try:
        books_resp = (
            client.table("bunny_books")
            .select("book_data")
            .eq("owner", owner)
            .execute()
        )
        progress_resp = (
            client.table("bunny_progress")
            .select("book_id,chapter_index")
            .eq("owner", owner)
            .execute()
        )
        bookmarks_resp = (
            client.table("bunny_bookmarks")
            .select("bookmark_data")
            .eq("owner", owner)
            .execute()
        )

        books = []
        for row in books_resp.data or []:
            book = row.get("book_data")
            if isinstance(book, dict) and book.get("book_id"):
                books.append(book)

        last_read = {}
        for row in progress_resp.data or []:
            book_id = row.get("book_id")
            if book_id:
                try:
                    last_read[book_id] = max(0, int(row.get("chapter_index", 0)))
                except (TypeError, ValueError):
                    last_read[book_id] = 0

        bookmarks = []
        for row in bookmarks_resp.data or []:
            item = row.get("bookmark_data")
            if isinstance(item, dict) and item.get("book_id") and item.get("chapter_id"):
                bookmarks.append(item)

        return {"books": books, "last_read": last_read, "bookmarks": bookmarks}, None
    except Exception as exc:
        return {"books": [], "last_read": {}, "bookmarks": []}, str(exc)


def persist_book(book):
    client = get_supabase_client()
    if client is None:
        return False, "Supabase is not configured yet. Add SUPABASE_URL and SUPABASE_KEY in Streamlit Secrets."

    try:
        client.table("bunny_books").upsert(
            {
                "owner": storage_owner(),
                "book_id": book["book_id"],
                "book_data": book,
            },
            on_conflict="owner,book_id",
        ).execute()
        return True, None
    except Exception as exc:
        return False, str(exc)


def persist_progress(book_id: str, chapter_index: int):
    client = get_supabase_client()
    if client is None:
        return False, "Supabase is not configured."

    try:
        client.table("bunny_progress").upsert(
            {
                "owner": storage_owner(),
                "book_id": book_id,
                "chapter_index": max(0, int(chapter_index)),
            },
            on_conflict="owner,book_id",
        ).execute()
        return True, None
    except Exception as exc:
        return False, str(exc)


def persist_bookmark(item):
    client = get_supabase_client()
    if client is None:
        return False, "Supabase is not configured."

    try:
        client.table("bunny_bookmarks").upsert(
            {
                "owner": storage_owner(),
                "book_id": item["book_id"],
                "chapter_id": item["chapter_id"],
                "bookmark_data": item,
            },
            on_conflict="owner,book_id,chapter_id",
        ).execute()
        return True, None
    except Exception as exc:
        return False, str(exc)


def delete_persistent_bookmark(book_id: str, chapter_id: str):
    client = get_supabase_client()
    if client is None:
        return False, "Supabase is not configured."

    try:
        (
            client.table("bunny_bookmarks")
            .delete()
            .eq("owner", storage_owner())
            .eq("book_id", book_id)
            .eq("chapter_id", chapter_id)
            .execute()
        )
        return True, None
    except Exception as exc:
        return False, str(exc)


# =========================================================
# STATE
# =========================================================
def initialize_state():
    defaults = {
        "page": "Home",
        "books": DEMO_BOOKS.copy(),
        "current_book_id": DEMO_BOOKS[0]["book_id"],
        "current_chapter_index": 0,
        "last_read": {DEMO_BOOKS[0]["book_id"]: 0},
        "bookmarks": [],
        "font_size": 19,
        "pending_import": None,
        "import_mode": "Paste / Type Text",
        "tts_nonce": 0,
        "reader_notice": "",
        "storage_loaded": False,
        "storage_error": "",
        "library_audience": "Mommy",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    builtin_ids = {book["book_id"] for book in DEMO_BOOKS}
    legacy_builtin_ids = {"technical-product-manager-ai-llm-chinese"}
    retired_builtin_ids = {
        "product-development-interview-prep",
        "problem-solving-series-i",
        "structural-thinking-for-reasoning-complexity-idea",
        "tell-me-about-yourself",
        "career-story-and-motivation",
        "polly-you-are-gorgeous",
    }

    # Load persistent user data once per browser session.
    if not st.session_state.storage_loaded:
        persistent, error = load_persistent_state()
        st.session_state.storage_loaded = True
        st.session_state.storage_error = error or ""

        persisted_books = [
            book for book in persistent.get("books", [])
            if book.get("book_id") not in builtin_ids
            and book.get("book_id") not in legacy_builtin_ids
            and book.get("book_id") not in retired_builtin_ids
        ]
        st.session_state.books = DEMO_BOOKS.copy() + persisted_books

        persisted_progress = persistent.get("last_read", {})
        if isinstance(persisted_progress, dict):
            st.session_state.last_read.update(persisted_progress)

        persisted_bookmarks = persistent.get("bookmarks", [])
        if isinstance(persisted_bookmarks, list):
            st.session_state.bookmarks = persisted_bookmarks

    # Keep built-in books synced to the latest code while preserving user books.
    current_books = st.session_state.get("books", [])
    imported_books = [
        book for book in current_books
        if book.get("book_id") not in builtin_ids
        and book.get("book_id") not in legacy_builtin_ids
        and book.get("book_id") not in retired_builtin_ids
    ]
    st.session_state.books = DEMO_BOOKS.copy() + imported_books

    # Older imported books may predate the Kids/Mommy split.
    # Treat them as Mommy unless they were explicitly saved to Kids.
    for stored_book in st.session_state.books:
        if stored_book.get("book_id") not in builtin_ids:
            stored_book.setdefault("audience", "Mommy")

    # Migrate an open legacy TPM book to the current built-in version.
    if st.session_state.get("current_book_id") in legacy_builtin_ids:
        st.session_state.current_book_id = "technical-product-manager-ai-llm"
        st.session_state.current_chapter_index = 0

    # If a removed built-in book was open in the current browser session,
    # safely move the reader to the first remaining built-in book.
    if st.session_state.get("current_book_id") in retired_builtin_ids:
        st.session_state.current_book_id = DEMO_BOOKS[0]["book_id"]
        st.session_state.current_chapter_index = 0

    old_tpm_progress = st.session_state.last_read.pop(
        "technical-product-manager-ai-llm-chinese", None
    )
    if old_tpm_progress is not None:
        st.session_state.last_read["technical-product-manager-ai-llm"] = old_tpm_progress

    for retired_book_id in retired_builtin_ids:
        st.session_state.last_read.pop(retired_book_id, None)

    st.session_state.bookmarks = [
        bookmark
        for bookmark in st.session_state.get("bookmarks", [])
        if bookmark.get("book_id") not in legacy_builtin_ids
        and bookmark.get("book_id") not in retired_builtin_ids
    ]

    for book in DEMO_BOOKS:
        st.session_state.last_read.setdefault(book["book_id"], 0)



# =========================================================
# HELPERS
# =========================================================
def slugify(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9\u0E00-\u0E7F\u4E00-\u9FFF]+", "-", value)
    return value.strip("-") or "book"


def get_book(book_id: str):
    return next((b for b in st.session_state.books if b["book_id"] == book_id), None)


def current_book():
    return get_book(st.session_state.current_book_id)


def book_audience(book) -> str:
    """Return the library audience while keeping older Supabase books compatible."""
    audience = str(book.get("audience", "")).strip()
    if audience in {"Kids", "Mommy"}:
        return audience

    category = str(book.get("category", "")).lower()
    if "children" in category or category == "family":
        return "Kids"
    return "Mommy"


def calculate_progress(book, chapter_index: int) -> int:
    total = max(len(book.get("chapters", [])), 1)
    current = min(max(chapter_index + 1, 1), total)
    return round((current / total) * 100)


def open_book(book_id: str, chapter_index=None):
    # Opening a book from the app should start from app state, not a stale deep-link.
    try:
        st.query_params.clear()
    except Exception:
        pass
    book = get_book(book_id)
    if not book:
        return
    st.session_state.current_book_id = book_id
    st.session_state.library_audience = book_audience(book)
    if chapter_index is None:
        chapter_index = st.session_state.last_read.get(book_id, 0)
    chapter_index = max(0, min(chapter_index, max(len(book["chapters"]) - 1, 0)))
    st.session_state.current_chapter_index = chapter_index
    st.session_state.last_read[book_id] = chapter_index
    ok, error = persist_progress(book_id, chapter_index)
    if not ok and storage_is_configured():
        st.session_state.storage_error = error or "Could not save reading progress."
    st.session_state.page = "Reader"


def open_chapter(index: int):
    book = current_book()
    if not book:
        return
    index = max(0, min(index, len(book["chapters"]) - 1))
    st.session_state.current_chapter_index = index
    st.session_state.last_read[book["book_id"]] = index
    ok, error = persist_progress(book["book_id"], index)
    if not ok and storage_is_configured():
        st.session_state.storage_error = error or "Could not save reading progress."


def is_bookmarked(book_id: str, chapter_id: str) -> bool:
    return any(
        x["book_id"] == book_id and x["chapter_id"] == chapter_id
        for x in st.session_state.bookmarks
    )


def toggle_bookmark(book, chapter):
    existing = [
        x for x in st.session_state.bookmarks
        if x["book_id"] == book["book_id"] and x["chapter_id"] == chapter["chapter_id"]
    ]

    if existing:
        ok, error = delete_persistent_bookmark(book["book_id"], chapter["chapter_id"])
        if ok:
            st.session_state.bookmarks = [
                x for x in st.session_state.bookmarks
                if not (
                    x["book_id"] == book["book_id"]
                    and x["chapter_id"] == chapter["chapter_id"]
                )
            ]
            st.session_state.reader_notice = "Bookmark removed."
        else:
            st.session_state.reader_notice = f"Could not remove bookmark: {error}"
    else:
        item = {
            "book_id": book["book_id"],
            "book_title": book["title"],
            "chapter_id": chapter["chapter_id"],
            "chapter_title": chapter["chapter_title"],
            "chapter_index": st.session_state.current_chapter_index,
            "saved_item": chapter["chapter_title"],
        }
        ok, error = persist_bookmark(item)
        if ok:
            st.session_state.bookmarks.append(item)
            st.session_state.reader_notice = "Saved to Bookmarks."
        else:
            st.session_state.reader_notice = f"Could not save bookmark: {error}"


def split_text_into_chapters(text: str, title: str):
    cleaned = re.sub(r"\r\n?", "\n", text).strip()
    if not cleaned:
        return []

    # Try to split on obvious chapter/day/section headings.
    heading_pattern = re.compile(
        r"(?im)^(?=(?:chapter|day|section|part|บทที่|วันที่)\s*[\w\dIVXivxก-๙-]*.*$)"
    )
    parts = [p.strip() for p in heading_pattern.split(cleaned) if p.strip()]

    # If no natural headings were found, split by paragraphs into readable chunks.
    if len(parts) <= 1:
        paragraphs = [p.strip() for p in re.split(r"\n{2,}", cleaned) if p.strip()]
        chunks = []
        bucket = []
        char_count = 0
        for paragraph in paragraphs:
            bucket.append(paragraph)
            char_count += len(paragraph)
            if char_count >= 3000:
                chunks.append("\n\n".join(bucket))
                bucket = []
                char_count = 0
        if bucket:
            chunks.append("\n\n".join(bucket))
        parts = chunks or [cleaned]

    chapters = []
    for idx, part in enumerate(parts, start=1):
        lines = [ln.strip() for ln in part.splitlines() if ln.strip()]
        candidate = lines[0] if lines else f"Chapter {idx}"
        chapter_title = candidate[:90] if len(candidate) <= 90 else f"Chapter {idx}"
        chapters.append(
            {
                "chapter_id": slugify(f"{title}-{idx}-{chapter_title}"),
                "chapter_title": chapter_title,
                "content": part,
                "order": idx,
            }
        )
    return chapters


def save_book(title, author, content_type, category, description, raw_text, audience="Mommy"):
    title = title.strip()
    raw_text = raw_text.strip()
    if not title or not raw_text:
        return False, "Please add both a title and readable content."

    base_id = slugify(title)
    book_id = base_id
    counter = 2
    existing_ids = {b["book_id"] for b in st.session_state.books}
    while book_id in existing_ids:
        book_id = f"{base_id}-{counter}"
        counter += 1

    chapters = split_text_into_chapters(raw_text, title)
    if not chapters:
        return False, "No readable content was found."

    book = {
        "book_id": book_id,
        "title": title,
        "subtitle": category.strip() or "Imported content",
        "author": author.strip() or "Unknown source",
        "content_type": content_type,
        "category": category.strip() or "General",
        "description": description.strip(),
        "cover_emoji": "📖",
        "audience": audience if audience in {"Kids", "Mommy"} else "Mommy",
        "chapters": chapters,
    }

    # Save to Supabase first so the UI never claims success for a session-only book.
    ok, error = persist_book(book)
    if not ok:
        return False, f"Could not save permanently: {error}"

    st.session_state.books.append(book)
    st.session_state.last_read[book_id] = 0
    persist_progress(book_id, 0)
    st.session_state.current_book_id = book_id
    st.session_state.current_chapter_index = 0
    st.session_state.pending_import = None
    return True, book_id


# =========================================================
# EXTRACTION
# =========================================================
def extract_txt(uploaded_file) -> str:
    raw = uploaded_file.read()
    for encoding in ("utf-8", "utf-8-sig", "cp874", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def extract_pdf(uploaded_file) -> str:
    reader = PdfReader(uploaded_file)
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n\n".join(pages).strip()


def extract_docx(uploaded_file) -> str:
    doc = Document(uploaded_file)
    return "\n\n".join(p.text for p in doc.paragraphs if p.text.strip()).strip()


def extract_url(url: str) -> str:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Linux; Android 13; Mobile) "
            "AppleWebKit/537.36 Chrome/120 Safari/537.36"
        )
    }
    response = requests.get(url, headers=headers, timeout=10)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    for tag in soup(["script", "style", "nav", "footer", "aside", "form", "noscript"]):
        tag.decompose()

    article = soup.find("article") or soup.find("main") or soup.body
    if article is None:
        return ""

    blocks = []
    for el in article.find_all(["h1", "h2", "h3", "p", "li"]):
        text = " ".join(el.get_text(" ", strip=True).split())
        if len(text) >= 20:
            blocks.append(text)

    return "\n\n".join(blocks).strip()


# =========================================================
# CSS
# =========================================================
def inject_css():
    st.markdown(
        f"""
        <style>
        :root {{
            --pearl: {CI["pearl"]};
            --rose: {CI["rose"]};
            --lavender: {CI["lavender"]};
            --blue: {CI["blue"]};
            --sage: {CI["sage"]};
            --gold: {CI["gold"]};
            --taupe: {CI["taupe"]};
            --brown: {CI["brown"]};
            --muted: {CI["muted"]};
            --white: {CI["white"]};
        }}

        html, body, [class*="css"], [class*="st-"] {{
            color: var(--taupe);
        }}

        html {{
            background: var(--pearl);
        }}

        body {{
            background: var(--pearl);
        }}

        * {{
            box-sizing: border-box !important;
        }}

        [data-testid="stAppViewContainer"] {{
            background:
                radial-gradient(circle at 8% 8%, rgba(232,183,200,.18), transparent 20%),
                radial-gradient(circle at 92% 15%, rgba(207,199,232,.17), transparent 23%),
                radial-gradient(circle at 12% 88%, rgba(201,214,193,.20), transparent 24%),
                radial-gradient(circle at 88% 92%, rgba(220,200,161,.17), transparent 22%),
                linear-gradient(180deg, #FFFDFC 0%, #FAF4EE 46%, #F7F1EA 100%);
            color: var(--taupe);
        }}

        [data-testid="stHeader"] {{
            display: none !important;
        }}

        [data-testid="stToolbar"] {{
            visibility: hidden;
        }}

        #MainMenu, footer {{
            visibility: hidden;
        }}

        .block-container {{
            width: 100%;
            max-width: 760px;
            padding-top: 1rem;
            padding-bottom: 6.2rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }}

        h1, h2, h3, h4, h5, h6 {{
            color: var(--brown) !important;
            font-family: Georgia, "Times New Roman", serif !important;
            letter-spacing: .01em;
        }}

        p, li, label, span, div {{
            color: var(--taupe);
        }}

        a {{
            color: var(--brown) !important;
        }}

        /* ---------- WATERCOLOR STORYBOOK ART ---------- */
        .st-key-storybook_hero {{
            margin: 2px 0 14px 0;
            border-radius: 22px;
            overflow: hidden;
            border: 1px solid rgba(220,200,161,.55);
            box-shadow: 0 9px 28px rgba(138,116,104,.10);
            background: rgba(255,255,255,.72);
        }}

        .st-key-storybook_hero [data-testid="stImage"],
        .st-key-storybook_hero img {{
            width: 100% !important;
            border-radius: 21px !important;
            display: block !important;
        }}

        .st-key-storybook_footer {{
            margin-top: 26px;
            padding: 4px 20px 12px;
            border-radius: 22px;
            background:
                linear-gradient(
                    180deg,
                    rgba(255,255,255,0) 0%,
                    rgba(255,250,246,.84) 34%,
                    rgba(247,241,234,.94) 100%
                );
        }}

        .st-key-storybook_footer [data-testid="stImage"] {{
            max-width: 420px;
            margin: 0 auto;
        }}

        .st-key-storybook_footer img {{
            border-radius: 18px !important;
            mix-blend-mode: multiply;
        }}

        .storybook-footer-words {{
            text-align: center;
            font-family: Georgia, "Times New Roman", serif;
            color: #A58A7D !important;
            font-size: 12px;
            letter-spacing: .28em;
            margin: 5px 0 4px;
        }}

        .st-key-reader_bunny_art {{
            margin: 6px 0 2px;
        }}

        .st-key-reader_bunny_art [data-testid="stImage"] {{
            border-radius: 18px;
            overflow: hidden;
            border: 1px solid rgba(232,183,200,.30);
            box-shadow: 0 6px 18px rgba(138,116,104,.07);
        }}

        .st-key-reader_bunny_art img {{
            border-radius: 18px !important;
            display: block !important;
        }}

        .soft-card {{
            background:
                linear-gradient(145deg, rgba(255,255,255,.96), rgba(255,249,246,.92));
            border: 1px solid #E9DDD7;
            border-radius: 20px;
            padding: 16px;
            margin-bottom: 12px;
            box-shadow: 0 6px 18px rgba(138,116,104,.07);
            overflow-wrap: anywhere;
        }}

        .book-card {{
            background:
                linear-gradient(145deg, rgba(255,255,255,.97), rgba(252,247,244,.94));
            border: 1px solid #E9DDD7;
            border-radius: 20px;
            padding: 14px;
            margin-bottom: 10px;
            box-shadow: 0 5px 16px rgba(138,116,104,.06);
        }}

        .book-title {{
            font-family: Georgia, "Times New Roman", serif;
            color: var(--brown) !important;
            font-size: 19px;
            line-height: 1.2;
            margin-bottom: 4px;
        }}

        .muted {{
            color: var(--muted) !important;
            font-size: 13px;
        }}

        .chip-row {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin: 8px 0;
        }}

        .chip {{
            border-radius: 999px;
            padding: 6px 10px;
            background: #FAEEF2;
            color: var(--brown) !important;
            font-size: 12px;
            border: 1px solid #EAD4DC;
        }}

        .progress-shell {{
            width: 100%;
            height: 8px;
            border-radius: 999px;
            background: #EEE7E2;
            overflow: hidden;
            margin-top: 8px;
        }}

        .progress-fill {{
            height: 100%;
            border-radius: 999px;
            background: linear-gradient(90deg, var(--rose), var(--lavender), var(--blue));
        }}

        .section-title {{
            font-family: Georgia, "Times New Roman", serif;
            color: var(--brown) !important;
            font-size: 21px;
            margin: 18px 0 8px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .section-title::before {{
            content: "❀";
            color: #CDA7B5 !important;
            font-size: 15px;
        }}

        .section-title::after {{
            content: "";
            height: 1px;
            flex: 1;
            background: linear-gradient(
                90deg,
                rgba(232,183,200,.55),
                rgba(207,199,232,.28),
                transparent
            );
        }}

        .reader-paper {{
            background: rgba(255,255,255,.94);
            border: 1px solid #E9DDD7;
            border-radius: 20px;
            padding: 18px 16px;
            margin: 8px 0 12px;
            box-shadow: 0 6px 20px rgba(138,116,104,.06);
        }}

        .reader-content {{
            white-space: pre-wrap;
            line-height: 1.78;
            overflow-wrap: anywhere;
            color: var(--taupe) !important;
        }}


        /* ---------- CHINESE VOCABULARY READER ---------- */
        .vocab-reader {{
            --reader-font-size: 18px;
            margin: 4px 0 10px;
        }}

        .vocab-intro {{
            margin: 0 0 9px;
            padding: 9px 11px;
            border-radius: 14px;
            background: #FFF9F5;
            border: 1px solid #EEE2DB;
            color: var(--muted) !important;
            font-size: calc(var(--reader-font-size) * .78);
            line-height: 1.42;
        }}

        .vocab-list {{
            display: flex;
            flex-direction: column;
            gap: 7px;
        }}

        .vocab-card {{
            position: relative;
            margin: 0;
            padding: 11px 12px 10px 13px;
            border-radius: 17px;
            border: 1px solid #EADFD9;
            background: rgba(255,255,255,.96);
            box-shadow: 0 4px 13px rgba(138,116,104,.045);
            overflow: hidden;
        }}

        .vocab-topline {{
            display: flex;
            align-items: flex-start;
            gap: 8px;
            min-width: 0;
        }}

        .vocab-number {{
            flex: 0 0 auto;
            min-width: 25px;
            height: 25px;
            padding: 0 6px;
            border-radius: 999px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            background: #F5E8ED;
            border: 1px solid #E7D1DA;
            color: var(--muted) !important;
            font-size: 11px;
            font-weight: 750;
            margin-top: 6px;
        }}

        .vocab-wordblock {{
            min-width: 0;
            flex: 1;
        }}

        .vocab-chinese {{
            display: inline-block;
            color: #B88FA6;
            font-family: "Noto Sans SC", "Microsoft YaHei", "PingFang SC", Arial, sans-serif;
            font-size: clamp(38px, calc(var(--reader-font-size) * 1.95), 56px);
            font-weight: 900;
            line-height: 1.04;
            letter-spacing: .01em;
            margin: 0;
        }}

        @supports ((-webkit-background-clip: text) or (background-clip: text)) {{
            .vocab-chinese {{
                background-image: var(--vocab-gradient);
                background-size: 100% 100%;
                background-clip: text;
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
            }}
        }}

        .vocab-pinyin {{
            margin-top: 2px;
            color: var(--muted) !important;
            font-size: calc(var(--reader-font-size) * .82);
            font-style: italic;
            line-height: 1.25;
        }}

        .vocab-meaning {{
            margin-top: 4px;
            color: var(--brown) !important;
            font-size: calc(var(--reader-font-size) * .98);
            font-weight: 700;
            line-height: 1.3;
        }}

        .vocab-roots {{
            margin-top: 7px;
            padding: 7px 9px;
            border-radius: 12px;
            background: linear-gradient(90deg, #FFF7F2, #FAF5FB);
            border: 1px solid #F0E4DE;
            color: var(--taupe) !important;
            font-size: calc(var(--reader-font-size) * .76);
            line-height: 1.42;
            overflow-wrap: anywhere;
        }}

        /* ---------- CHAPTER / SUB-CATEGORY NAVIGATOR ---------- */
        .st-key-reader_chapter_nav {{
            margin-top: 0 !important;
            padding: 0 1px !important;
            min-width: 0 !important;
            overflow: hidden !important;
        }}

        .st-key-reader_chapter_nav [role="radiogroup"] {{
            display: flex !important;
            flex-direction: row !important;
            flex-wrap: nowrap !important;
            gap: 6px !important;
            overflow-x: auto !important;
            overflow-y: hidden !important;
            width: 100% !important;
            padding: 1px 1px 5px !important;
            scrollbar-width: none !important;
            -webkit-overflow-scrolling: touch;
            scroll-snap-type: x proximity;
        }}

        .st-key-reader_chapter_nav [role="radiogroup"]::-webkit-scrollbar {{
            display: none !important;
        }}

        .st-key-reader_chapter_nav [role="radiogroup"] > label > div:first-child,
        .st-key-reader_chapter_nav label[data-baseweb="radio"] > div:first-child,
        .st-key-reader_chapter_nav input[type="radio"] {{
            display: none !important;
        }}

        .st-key-reader_chapter_nav [role="radiogroup"] > label {{
            flex: 0 0 auto !important;
            min-height: 31px !important;
            margin: 0 !important;
            padding: 5px 10px !important;
            border-radius: 999px !important;
            border: 1px solid #E6D8D1 !important;
            background: #FAEEF2 !important;
            color: var(--taupe) !important;
            scroll-snap-align: start;
            cursor: pointer;
            white-space: nowrap !important;
        }}

        .st-key-reader_chapter_nav [role="radiogroup"] > label:nth-child(6n+2) {{ background: #EEF4EA !important; }}
        .st-key-reader_chapter_nav [role="radiogroup"] > label:nth-child(6n+3) {{ background: #F6ECE3 !important; }}
        .st-key-reader_chapter_nav [role="radiogroup"] > label:nth-child(6n+4) {{ background: #F0ECFA !important; }}
        .st-key-reader_chapter_nav [role="radiogroup"] > label:nth-child(6n+5) {{ background: #EAF3F8 !important; }}
        .st-key-reader_chapter_nav [role="radiogroup"] > label:nth-child(6n+6) {{ background: #FBF3DF !important; }}

        .st-key-reader_chapter_nav [role="radiogroup"] > label:has(input:checked) {{
            background: #F8E3EB !important;
            border-color: #D8AABD !important;
            box-shadow: inset 0 0 0 1px rgba(216,170,189,.34) !important;
            font-weight: 750 !important;
        }}

        .st-key-reader_chapter_nav [role="radiogroup"] > label p,
        .st-key-reader_chapter_nav [role="radiogroup"] > label span {{
            white-space: nowrap !important;
            font-size: 11.5px !important;
            color: var(--taupe) !important;
        }}

        /*
           Streamlit Cloud can render keyed containers with a slightly different
           DOM/class structure from local Streamlit. This fallback intentionally
           targets the app's only radio control directly so the chapter navigator
           keeps the same pastel pill layout in both environments.
        */
        [data-testid="stRadio"] {{
            width: 100% !important;
            min-width: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
            overflow: hidden !important;
        }}

        [data-testid="stRadio"] [role="radiogroup"] {{
            display: flex !important;
            flex-direction: row !important;
            flex-wrap: nowrap !important;
            align-items: center !important;
            gap: 6px !important;
            width: 100% !important;
            min-width: 0 !important;
            overflow-x: auto !important;
            overflow-y: hidden !important;
            padding: 1px 1px 5px !important;
            scrollbar-width: none !important;
            -webkit-overflow-scrolling: touch !important;
            scroll-snap-type: x proximity;
        }}

        [data-testid="stRadio"] [role="radiogroup"]::-webkit-scrollbar {{
            display: none !important;
        }}

        [data-testid="stRadio"] [role="radiogroup"] > label,
        [data-testid="stRadio"] label[data-baseweb="radio"] {{
            flex: 0 0 auto !important;
            width: auto !important;
            min-width: max-content !important;
            max-width: none !important;
            min-height: 31px !important;
            margin: 0 !important;
            padding: 5px 10px !important;
            border-radius: 999px !important;
            border: 1px solid #E6D8D1 !important;
            background: #FAEEF2 !important;
            color: var(--taupe) !important;
            cursor: pointer !important;
            white-space: nowrap !important;
            scroll-snap-align: start;
            box-sizing: border-box !important;
        }}

        [data-testid="stRadio"] [role="radiogroup"] > label:nth-child(6n+2),
        [data-testid="stRadio"] label[data-baseweb="radio"]:nth-child(6n+2) {{ background: #EEF4EA !important; }}
        [data-testid="stRadio"] [role="radiogroup"] > label:nth-child(6n+3),
        [data-testid="stRadio"] label[data-baseweb="radio"]:nth-child(6n+3) {{ background: #F6ECE3 !important; }}
        [data-testid="stRadio"] [role="radiogroup"] > label:nth-child(6n+4),
        [data-testid="stRadio"] label[data-baseweb="radio"]:nth-child(6n+4) {{ background: #F0ECFA !important; }}
        [data-testid="stRadio"] [role="radiogroup"] > label:nth-child(6n+5),
        [data-testid="stRadio"] label[data-baseweb="radio"]:nth-child(6n+5) {{ background: #EAF3F8 !important; }}
        [data-testid="stRadio"] [role="radiogroup"] > label:nth-child(6n+6),
        [data-testid="stRadio"] label[data-baseweb="radio"]:nth-child(6n+6) {{ background: #FBF3DF !important; }}

        [data-testid="stRadio"] [role="radiogroup"] > label:has(input:checked),
        [data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked),
        [data-testid="stRadio"] [role="radiogroup"] > label:has([aria-checked="true"]),
        [data-testid="stRadio"] label[data-baseweb="radio"]:has([aria-checked="true"]) {{
            background: #F8E3EB !important;
            border-color: #D8AABD !important;
            box-shadow: inset 0 0 0 1px rgba(216,170,189,.34) !important;
            font-weight: 750 !important;
        }}

        [data-testid="stRadio"] [role="radiogroup"] > label p,
        [data-testid="stRadio"] [role="radiogroup"] > label span,
        [data-testid="stRadio"] label[data-baseweb="radio"] p,
        [data-testid="stRadio"] label[data-baseweb="radio"] span {{
            white-space: nowrap !important;
            overflow-wrap: normal !important;
            word-break: keep-all !important;
            font-size: 11.5px !important;
            color: var(--taupe) !important;
            margin: 0 !important;
        }}

        /* Stable HTML chapter navigator: independent of Streamlit widget DOM. */
        .chapter-nav-scroll {{
            display: flex !important;
            flex-direction: row !important;
            flex-wrap: nowrap !important;
            align-items: center !important;
            gap: 7px !important;
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: auto !important;
            overflow-y: hidden !important;
            padding: 2px 1px 7px !important;
            margin: 0 !important;
            box-sizing: border-box !important;
            scrollbar-width: none !important;
            -webkit-overflow-scrolling: touch !important;
            scroll-snap-type: x proximity;
        }}

        .chapter-nav-scroll::-webkit-scrollbar {{
            display: none !important;
        }}

        .chapter-nav-pill {{
            flex: 0 0 auto !important;
            display: inline-flex !important;
            align-items: center !important;
            gap: 7px !important;
            min-height: 34px !important;
            padding: 6px 11px !important;
            border-radius: 999px !important;
            border: 1px solid #E6D8D1 !important;
            background: #FAEEF2 !important;
            color: var(--taupe) !important;
            text-decoration: none !important;
            white-space: nowrap !important;
            word-break: normal !important;
            overflow-wrap: normal !important;
            font-size: 12px !important;
            line-height: 1.15 !important;
            scroll-snap-align: start;
            box-sizing: border-box !important;
        }}

        .chapter-nav-pill:nth-child(6n+2) {{ background: #EEF4EA !important; }}
        .chapter-nav-pill:nth-child(6n+3) {{ background: #F6ECE3 !important; }}
        .chapter-nav-pill:nth-child(6n+4) {{ background: #F0ECFA !important; }}
        .chapter-nav-pill:nth-child(6n+5) {{ background: #EAF3F8 !important; }}
        .chapter-nav-pill:nth-child(6n+6) {{ background: #FBF3DF !important; }}

        .chapter-nav-pill.active {{
            background: #F8E3EB !important;
            border-color: #D8AABD !important;
            box-shadow: inset 0 0 0 1px rgba(216,170,189,.34) !important;
            font-weight: 750 !important;
        }}

        .chapter-nav-dot {{
            width: 14px !important;
            height: 14px !important;
            flex: 0 0 14px !important;
            border-radius: 50% !important;
            background: #FFFFFF !important;
            border: 1.5px solid #C9B8AF !important;
            box-sizing: border-box !important;
        }}

        .chapter-nav-pill.active .chapter-nav-dot {{
            background: #F56B73 !important;
            border: 4px solid #F56B73 !important;
            box-shadow: inset 0 0 0 2px #FFFFFF !important;
        }}

        .empty-state {{
            text-align: center;
            background: rgba(255,255,255,.88);
            border: 1px dashed #D9C7BD;
            border-radius: 20px;
            padding: 28px 16px;
            margin: 16px 0;
        }}

        .pastel-note {{
            border-radius: 16px;
            padding: 12px 14px;
            background: #F8EEF1;
            border: 1px solid #E7CDD6;
            color: var(--taupe) !important;
            margin: 10px 0;
        }}

        .pastel-success {{
            background: #EEF3EA;
            border: 1px solid var(--sage);
        }}

        .pastel-info {{
            background: #EDF5FA;
            border: 1px solid var(--blue);
        }}

        .pastel-warning {{
            background: #FBF4E8;
            border: 1px solid var(--gold);
        }}

        .decor-line {{
            text-align: center;
            letter-spacing: .45em;
            color: var(--muted) !important;
            margin: 10px 0 4px;
        }}

        /* ---------- STREAMLIT BUTTONS ---------- */
        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stBaseButton-secondary"],
        [data-testid="stBaseButton-primary"] {{
            min-height: 44px !important;
            border-radius: 14px !important;
            border: 1px solid #DCC8D0 !important;
            background: var(--rose) !important;
            color: var(--taupe) !important;
            box-shadow: none !important;
            font-weight: 650 !important;
            transition: all .15s ease !important;
            width: 100% !important;
        }}

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        [data-testid="stBaseButton-secondary"]:hover,
        [data-testid="stBaseButton-primary"]:hover {{
            background: #DFABB8 !important;
            color: var(--taupe) !important;
            border-color: #D2A2AF !important;
        }}

        .stButton > button:active,
        .stDownloadButton > button:active,
        [data-testid="stBaseButton-secondary"]:active,
        [data-testid="stBaseButton-primary"]:active {{
            background: var(--lavender) !important;
            color: var(--taupe) !important;
            transform: translateY(1px);
        }}

        .stButton > button:focus,
        .stButton > button:focus-visible,
        .stDownloadButton > button:focus,
        [data-testid="stBaseButton-secondary"]:focus,
        [data-testid="stBaseButton-primary"]:focus {{
            outline: 3px solid rgba(220,200,161,.68) !important;
            outline-offset: 2px !important;
            box-shadow: none !important;
            color: var(--taupe) !important;
        }}

        .stButton > button p,
        .stButton > button span,
        .stDownloadButton > button p,
        .stDownloadButton > button span {{
            color: var(--taupe) !important;
        }}


        .st-key-reader_top_actions .stButton > button:active,
        .st-key-reader_top_actions .stButton > button:focus,
        .st-key-reader_top_actions .stButton > button:focus-visible {{
            color: var(--taupe) !important;
            outline: 3px solid rgba(220,200,161,.68) !important;
            outline-offset: 2px !important;
            box-shadow: none !important;
        }}

        .st-key-reader_top_actions .stButton > button:disabled {{
            opacity: .48 !important;
            color: var(--muted) !important;
            filter: saturate(.65) !important;
        }}

        .st-key-reader_top_actions .stButton > button p,
        .st-key-reader_top_actions .stButton > button span {{
            color: var(--taupe) !important;
        }}

        /* ---------- INPUTS ---------- */
        input,
        textarea,
        select,
        [data-baseweb="input"] input,
        [data-baseweb="textarea"] textarea {{
            color: var(--taupe) !important;
            background: var(--white) !important;
            border-color: #DCCFC7 !important;
            caret-color: var(--brown) !important;
            font-size: 16px !important;
        }}

        input:focus,
        textarea:focus,
        select:focus {{
            outline: 2px solid var(--gold) !important;
            outline-offset: 1px !important;
            box-shadow: none !important;
        }}

        [data-baseweb="select"] > div {{
            background: var(--white) !important;
            color: var(--taupe) !important;
            border-color: #DCCFC7 !important;
        }}

        [data-baseweb="popover"] {{
            color: var(--taupe) !important;
        }}

        [role="option"] {{
            color: var(--taupe) !important;
            background: var(--white) !important;
        }}

        [role="option"]:hover,
        [role="option"][aria-selected="true"] {{
            color: var(--taupe) !important;
            background: #F7E7ED !important;
        }}

        /* ---------- FILE UPLOADER ---------- */
        [data-testid="stFileUploader"] {{
            background: rgba(255,255,255,.9);
            border-radius: 16px;
        }}

        [data-testid="stFileUploaderDropzone"] {{
            background: #FFFDFB !important;
            border: 1px dashed #D7C8BF !important;
            color: var(--taupe) !important;
        }}

        [data-testid="stFileUploaderDropzone"] button {{
            background: var(--lavender) !important;
            color: var(--taupe) !important;
            border: 1px solid #BDB4D6 !important;
        }}

        /* ---------- CHECKS / RADIO ---------- */
        input[type="checkbox"],
        input[type="radio"] {{
            accent-color: var(--rose) !important;
        }}

        [role="checkbox"] svg,
        [role="radio"] svg {{
            color: var(--brown) !important;
        }}

        /* ---------- PROGRESS ---------- */
        [data-testid="stProgressBar"] > div > div {{
            background: linear-gradient(
                90deg,
                var(--rose),
                var(--lavender),
                var(--blue)
            ) !important;
        }}

        /* ---------- TABS ---------- */
        button[data-baseweb="tab"] {{
            color: var(--taupe) !important;
        }}

        button[data-baseweb="tab"][aria-selected="true"] {{
            color: var(--brown) !important;
            background: #F8E9EE !important;
        }}

        /* ---------- ALERTS ---------- */
        [data-testid="stAlert"] {{
            background: #FBF4E8 !important;
            color: var(--taupe) !important;
            border: 1px solid var(--gold) !important;
            border-radius: 14px !important;
        }}

        /* ---------- MOBILE NAV ---------- */
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] {{
            display: grid !important;
            grid-template-columns: repeat(5, minmax(0, 1fr)) !important;
            gap: 6px !important;
            width: 100% !important;
        }}

        .st-key-bunny_navigation [data-testid="stColumn"],
        .st-key-bunny_navigation [data-testid="column"] {{
            width: 100% !important;
            min-width: 0 !important;
        }}

        .st-key-bunny_navigation .stButton > button {{
            width: 100% !important;
            min-height: 46px !important;
            padding: 8px 2px !important;
        }}

        .st-key-bunny_navigation .stButton > button p {{
            font-size: clamp(9.5px, 2.7vw, 13px) !important;
            white-space: nowrap !important;
        }}

        /* ---------- LIBRARY FILTERS: 4 BUTTONS, ONE ROW ---------- */
        .st-key-library_filters [data-testid="stHorizontalBlock"] {{
            display: grid !important;
            grid-template-columns: repeat(4, minmax(0, 1fr)) !important;
            gap: 6px !important;
            width: 100% !important;
        }}

        .st-key-library_filters [data-testid="stColumn"],
        .st-key-library_filters [data-testid="column"] {{
            width: 100% !important;
            min-width: 0 !important;
        }}

        .st-key-library_filters .stButton > button {{
            width: 100% !important;
            min-height: 46px !important;
            padding: 8px 2px !important;
        }}

        .st-key-library_filters .stButton > button p {{
            font-size: clamp(11px, 3.2vw, 14px) !important;
            white-space: nowrap !important;
        }}

        @media (max-width: 520px) {{
            .st-key-storybook_hero {{
                border-radius: 17px;
                margin-top: 2px;
            }}

            .st-key-storybook_hero img {{
                border-radius: 16px !important;
            }}

            .st-key-storybook_footer {{
                padding-left: 8px;
                padding-right: 8px;
            }}

            .storybook-footer-words {{
                letter-spacing: .18em;
                font-size: 10px;
            }}

            .st-key-reader_bunny_art [data-testid="column"]:first-child,
            .st-key-reader_bunny_art [data-testid="column"]:last-child {{
                display: block !important;
            }}

            .block-container {{
                max-width: 100%;
                padding-left: 14px;
                padding-right: 14px;
                padding-top: .65rem;
                padding-bottom: 6.5rem;
            }}

            .hero-row {{
                align-items: center;
            }}

            .hero-bunny {{
                min-width: 72px;
                height: 72px;
                font-size: 40px;
            }}

            .hero-title {{
                font-size: 30px;
            }}

            .stButton > button {{
                min-height: 46px !important;
                padding-left: 8px !important;
                padding-right: 8px !important;
                font-size: 14px !important;
            }}

            [data-testid="column"] {{
                min-width: 0 !important;
            }}
        }}

        /* Keep reader controls on three compact rows on mobile. */
        .st-key-reader_heading_controls [data-testid="stHorizontalBlock"],
        .st-key-reader_top_actions [data-testid="stHorizontalBlock"] {{
            display: grid !important;
            gap: 6px !important;
            width: 100% !important;
        }}
        .st-key-reader_heading_controls [data-testid="stHorizontalBlock"] {{
            grid-template-columns: minmax(0, 2fr) repeat(2, minmax(0, .7fr)) !important;
        }}
        .st-key-reader_top_actions [data-testid="stHorizontalBlock"] {{
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
        }}
        .st-key-reader_heading_controls [data-testid="stColumn"],
        .st-key-reader_heading_controls [data-testid="column"],
        .st-key-reader_top_actions [data-testid="stColumn"],
        .st-key-reader_top_actions [data-testid="column"] {{
            width: 100% !important;
            min-width: 0 !important;
        }}
        .st-key-reader_top_actions .stButton > button p {{
            white-space: nowrap !important;
        }}
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(1) .stButton > button,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(1) .stButton > button:hover,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(1) .stButton > button:focus,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(1) .stButton > button:active {{
            background: #E8C8D2 !important;
            border-color: #E8C8D2 !important;
            color: var(--taupe) !important;
        }}
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(2) .stButton > button,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(2) .stButton > button:hover,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(2) .stButton > button:focus,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(2) .stButton > button:active {{
            background: #D3DDC9 !important;
            border-color: #D3DDC9 !important;
            color: var(--taupe) !important;
        }}
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(3) .stButton > button,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(3) .stButton > button:hover,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(3) .stButton > button:focus,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(3) .stButton > button:active {{
            background: #DFCDBD !important;
            border-color: #DFCDBD !important;
            color: var(--taupe) !important;
        }}
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(4) .stButton > button,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(4) .stButton > button:hover,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(4) .stButton > button:focus,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(4) .stButton > button:active {{
            background: #EAD8CF !important;
            border-color: #EAD8CF !important;
            color: var(--taupe) !important;
        }}
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(5) .stButton > button,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(5) .stButton > button:hover,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(5) .stButton > button:focus,
        .st-key-bunny_navigation [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"], [data-testid="column"]):nth-child(5) .stButton > button:active {{
            background: #DFD8F3 !important;
            border-color: #DFD8F3 !important;
            color: var(--taupe) !important;
        }}
        /* Fixed controls, with matching space above the scrollable content. */
        .st-key-bunny_navigation,
        .st-key-reader_fixed_audio {{
            position: fixed !important;
            left: 50% !important;
            transform: translateX(-50%);
            width: min(728px, calc(100% - 28px)) !important;
            max-width: 728px !important;
            background: #FFFDFC !important;
            box-sizing: border-box !important;
        }}
        .st-key-bunny_navigation {{
            top: 0 !important;
            height: calc(64px + env(safe-area-inset-top, 0px)) !important;
            padding: calc(8px + env(safe-area-inset-top, 0px)) 0 8px !important;
            z-index: 1001 !important;
        }}
        .st-key-bunny_navigation .stButton > button {{
            height: 48px !important;
        }}
        .st-key-reader_fixed_audio {{
            top: calc(64px + env(safe-area-inset-top, 0px)) !important;
            height: 128px !important;
            padding: 3px 0 2px !important;
            z-index: 1000 !important;
            overflow: hidden !important;
        }}
        .block-container {{
            padding-top: calc(80px + env(safe-area-inset-top, 0px)) !important;
        }}
        .block-container:has(.st-key-reader_fixed_audio) {{
            padding-top: calc(204px + env(safe-area-inset-top, 0px)) !important;
        }}
        [data-testid="stMain"] {{
            scroll-padding-top: calc(204px + env(safe-area-inset-top, 0px));
        }}

        /* A−, A+, and Saved share the left 1/4. Audio takes the right 3/4. */
        .st-key-reader_audio_control_row [data-testid="stHorizontalBlock"] {{
            display: grid !important;
            grid-template-columns: repeat(3, minmax(0, 1fr)) minmax(0, 9fr) !important;
            gap: 5px !important;
            width: 100% !important;
            align-items: center !important;
        }}

        .st-key-reader_audio_control_row [data-testid="stColumn"],
        .st-key-reader_audio_control_row [data-testid="column"] {{
            width: 100% !important;
            min-width: 0 !important;
        }}

        .st-key-reader_audio_control_row .stButton > button {{
            width: 100% !important;
            min-width: 0 !important;
            min-height: 40px !important;
            height: 40px !important;
            padding-left: 0 !important;
            padding-right: 0 !important;
            font-size: 12px !important;
        }}

        .st-key-reader_audio_control_row .stButton > button p {{
            white-space: nowrap !important;
            font-size: 12px !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# NAVIGATION
# =========================================================
def render_navigation():
    with st.container(key="bunny_navigation"):
        cols = st.columns(5, gap="small")
        items = [
            ("Home", "⌂ Home"),
            ("Kids", "🐰 Kids"),
            ("Mommy", "🌷 Mommy"),
            ("Add", "＋ Add"),
            ("Bookmarks", "♡ Saved"),
        ]
        for col, (page, label) in zip(cols, items):
            with col:
                if st.button(label, key=f"nav_{page}", use_container_width=True):
                    try:
                        st.query_params.clear()
                    except Exception:
                        pass
                    if page in {"Kids", "Mommy"}:
                        st.session_state.library_audience = page
                    st.session_state.page = page
                    st.rerun()


# =========================================================
# SHARED UI
# =========================================================
def render_hero():
    with st.container(key="storybook_hero"):
        st.image(
            str(ASSET_DIR / "bunny_header.jpg"),
            use_container_width=True,
        )


def render_storybook_footer():
    with st.container(key="storybook_footer"):
        st.image(
            str(ASSET_DIR / "floral_books_footer.jpg"),
            use_container_width=True,
        )
        st.markdown(
            '<div class="storybook-footer-words">'
            'READ ✦ DREAM ✦ DISCOVER ✦ GROW'
            '</div>',
            unsafe_allow_html=True,
        )


def render_progress(progress: int):
    progress = min(max(progress, 0), 100)
    st.markdown(
        f"""
        <div class="progress-shell" aria-label="Reading progress {progress}%">
            <div class="progress-fill" style="width:{progress}%"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_book_summary(book, chapter_index=0):
    total = max(len(book["chapters"]), 1)
    chapter_index = max(0, min(chapter_index, total - 1))
    progress = calculate_progress(book, chapter_index)
    chapter_title = book["chapters"][chapter_index]["chapter_title"] if book["chapters"] else "No chapters"

    st.markdown(
        f"""
        <div class="book-card">
            <div style="display:flex; gap:12px; align-items:flex-start;">
                <div style="
                    width:58px;height:78px;border-radius:12px;
                    background:linear-gradient(160deg,#F7E6EC,#EEE8F7);
                    border:1px solid #E6D7D0;
                    display:flex;align-items:center;justify-content:center;
                    font-size:30px;flex:0 0 auto;">
                    {escape(book.get("cover_emoji","📖"))}
                </div>
                <div style="min-width:0;flex:1;">
                    <div class="book-title">{escape(book["title"])}</div>
                    <div class="muted">{escape(book.get("author",""))}</div>
                    <div class="muted" style="margin-top:6px;">
                        {escape(chapter_title)} · {progress}%
                    </div>
                    <div class="progress-shell">
                        <div class="progress-fill" style="width:{progress}%"></div>
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# HOME
# =========================================================
def render_home():
    render_hero()

    st.markdown('<div class="section-title">Continue Reading</div>', unsafe_allow_html=True)
    book = current_book() or (st.session_state.books[0] if st.session_state.books else None)

    if book:
        chapter_index = st.session_state.last_read.get(book["book_id"], 0)
        render_book_summary(book, chapter_index)
        if st.button("Continue Reading", key="home_continue"):
            open_book(book["book_id"], chapter_index)
            st.rerun()
    else:
        st.markdown(
            """
            <div class="empty-state">
                <div style="font-size:48px;">🐰📚</div>
                <div class="book-title">Your library is waiting for its first story.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown('<div class="section-title">Choose a Library</div>', unsafe_allow_html=True)
    library_cols = st.columns(2, gap="small")

    with library_cols[0]:
        st.markdown(
            """
            <div class="soft-card" style="text-align:center;">
                <div style="font-size:38px;">🐰</div>
                <div class="book-title">Kids</div>
                <div class="muted">Stories, bedtime reading, and growing-up books.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Open Kids", key="home_kids", use_container_width=True):
            st.session_state.library_audience = "Kids"
            st.session_state.page = "Kids"
            st.rerun()

    with library_cols[1]:
        st.markdown(
            """
            <div class="soft-card" style="text-align:center;">
                <div style="font-size:38px;">🌷</div>
                <div class="book-title">Mommy</div>
                <div class="muted">TPM, Chinese, thinking, speaking, and personal learning.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Open Mommy", key="home_mommy", use_container_width=True):
            st.session_state.library_audience = "Mommy"
            st.session_state.page = "Mommy"
            st.rerun()

    st.markdown('<div class="section-title">Add Content</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="soft-card">
            <div style="font-size:34px;">🌷📖</div>
            <div class="book-title">Bring your next read into Bunny Reading</div>
            <div class="muted">Paste a link, upload a file, or add your own text.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("＋ Add Content", key="home_add"):
        st.session_state.page = "Add"
        st.rerun()


# =========================================================
# LIBRARY
# =========================================================
def render_library(audience=None):
    if audience not in {"Kids", "Mommy"}:
        audience = st.session_state.get("library_audience", "Mommy")
    st.session_state.library_audience = audience

    icon = "🐰" if audience == "Kids" else "🌷"
    st.markdown(f"## {icon} {audience}")

    search = st.text_input(
        "Search",
        placeholder=f"Search {audience} library...",
        label_visibility="collapsed",
        key=f"library_search_{audience}",
    )

    filters = ["All", "Books", "Articles", "Notes"]
    filter_key = f"library_filter_{audience}"
    if filter_key not in st.session_state:
        st.session_state[filter_key] = "All"

    with st.container(key="library_filters"):
        filter_cols = st.columns(4, gap="small")
        for col, label in zip(filter_cols, filters):
            with col:
                if st.button(label, key=f"filter_{audience}_{label}", use_container_width=True):
                    st.session_state[filter_key] = label
                    st.rerun()

    query = search.strip().lower()
    filtered = []
    for book in st.session_state.books:
        if book_audience(book) != audience:
            continue

        if query and query not in (
            f"{book.get('title','')} {book.get('author','')} "
            f"{book.get('category','')} {book.get('description','')}"
        ).lower():
            continue

        selected = st.session_state[filter_key]
        ctype = book.get("content_type", "Book").lower()

        if selected == "Books" and ctype not in {"book", "pdf", "docx", "txt"}:
            continue
        if selected == "Articles" and ctype not in {"article", "url"}:
            continue
        if selected == "Notes" and ctype not in {"note", "text"}:
            continue

        filtered.append(book)

    if not filtered:
        st.markdown(
            f"""
            <div class="empty-state">
                <div style="font-size:44px;">{icon}🌿</div>
                <div class="book-title">No matching books yet.</div>
                <div class="muted">Try another search or add new content to {audience}.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    for idx, book in enumerate(filtered):
        last = st.session_state.last_read.get(book["book_id"], 0)
        render_book_summary(book, last)
        if st.button("Continue Reading", key=f"lib_open_{audience}_{idx}_{book['book_id']}"):
            open_book(book["book_id"], last)
            st.rerun()


# =========================================================
# ADD CONTENT / IMPORT
# =========================================================
def render_add_content():
    if not storage_is_configured():
        st.markdown(
            """
            <div class="pastel-note pastel-warning">
                Permanent saving is not configured yet. Add your Supabase secrets before saving new books.
            </div>
            """,
            unsafe_allow_html=True,
        )
    elif st.session_state.get("storage_error"):
        st.markdown(
            f'<div class="pastel-note pastel-warning">Storage warning: {escape(st.session_state.storage_error)}</div>',
            unsafe_allow_html=True,
        )

    st.markdown("## Add Content")
    st.caption("Bring your favorite reads together in one cozy place.")

    modes = [
        ("Paste Link", "🔗", "#F8E6EC"),
        ("Upload PDF", "📄", "#EEEAF8"),
        ("Upload DOCX", "📝", "#EDF3E9"),
        ("Upload TXT", "📃", "#EAF3F8"),
        ("Paste / Type Text", "✎", "#F9F1E3"),
    ]

    for idx, (mode, icon, bg) in enumerate(modes):
        st.markdown(
            f"""
            <div class="soft-card" style="background:{bg};">
                <div style="display:flex;gap:12px;align-items:center;">
                    <div style="font-size:28px;">{icon}</div>
                    <div style="min-width:0;">
                        <div class="book-title">{escape(mode)}</div>
                        <div class="muted">
                            {
                                "Add from a webpage or article"
                                if mode == "Paste Link"
                                else "Choose a file from your device"
                                if mode.startswith("Upload")
                                else "Write or paste your own content"
                            }
                        </div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button(f"Choose {mode}", key=f"choose_mode_{idx}"):
            st.session_state.import_mode = mode
            st.session_state.pending_import = None
            st.rerun()

    st.markdown("---")
    st.markdown(f"### {st.session_state.import_mode}")

    mode = st.session_state.import_mode

    try:
        if mode == "Paste Link":
            url = st.text_input(
                "Website link",
                placeholder="https://example.com/article",
            )
            if st.button("Load Link", key="load_link"):
                if not url.strip():
                    st.warning("Please paste a link first.")
                else:
                    with st.spinner("Reading the page..."):
                        text = extract_url(url.strip())
                    if not text:
                        st.warning("I could not find readable article text on that page.")
                    else:
                        st.session_state.pending_import = {
                            "text": text,
                            "content_type": "Article",
                            "source": url.strip(),
                        }
                        st.rerun()

        elif mode == "Upload PDF":
            uploaded = st.file_uploader("Choose a PDF", type=["pdf"])
            if uploaded and st.button("Load PDF", key="load_pdf"):
                with st.spinner("Reading PDF..."):
                    text = extract_pdf(uploaded)
                st.session_state.pending_import = {
                    "text": text,
                    "content_type": "PDF",
                    "source": uploaded.name,
                }
                st.rerun()

        elif mode == "Upload DOCX":
            uploaded = st.file_uploader("Choose a DOCX", type=["docx"])
            if uploaded and st.button("Load DOCX", key="load_docx"):
                text = extract_docx(uploaded)
                st.session_state.pending_import = {
                    "text": text,
                    "content_type": "DOCX",
                    "source": uploaded.name,
                }
                st.rerun()

        elif mode == "Upload TXT":
            uploaded = st.file_uploader("Choose a TXT", type=["txt"])
            if uploaded and st.button("Load TXT", key="load_txt"):
                text = extract_txt(uploaded)
                st.session_state.pending_import = {
                    "text": text,
                    "content_type": "TXT",
                    "source": uploaded.name,
                }
                st.rerun()

        else:
            typed = st.text_area(
                "Paste or type your content",
                height=220,
                placeholder="Paste your reading material here...",
            )
            if st.button("Preview Text", key="preview_text"):
                if typed.strip():
                    st.session_state.pending_import = {
                        "text": typed.strip(),
                        "content_type": "Note",
                        "source": "Typed / pasted text",
                    }
                    st.rerun()
                else:
                    st.warning("Please add some text first.")

    except Exception as exc:
        st.markdown(
            f"""
            <div class="pastel-note">
                I couldn't import that content.<br>
                <span class="muted">{escape(str(exc))}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if st.session_state.pending_import:
        render_import_preview()


def render_import_preview():
    data = st.session_state.pending_import
    text = data.get("text", "")
    preview = text[:1800] + ("…" if len(text) > 1800 else "")

    st.markdown("### Import Preview")
    st.markdown(
        f"""
        <div class="soft-card">
            <div class="chip-row">
                <span class="chip">{escape(data.get("content_type","Content"))}</span>
                <span class="chip">{len(text):,} characters</span>
            </div>
            <div class="muted" style="margin-bottom:10px;">
                Source: {escape(data.get("source",""))}
            </div>
            <div style="white-space:pre-wrap;line-height:1.6;">
                {escape(preview)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.form("save_import_form"):
        title = st.text_input("Title")
        author = st.text_input("Author / Source", value=data.get("source", ""))
        audience = st.selectbox(
            "Library",
            ["Mommy", "Kids"],
            index=0,
            help="Choose where this book should appear.",
        )
        category = st.text_input("Category", value="Imported")
        description = st.text_area("Optional description", height=90)
        save = st.form_submit_button("Save to Library")

    if save:
        ok, result = save_book(
            title=title,
            author=author,
            content_type=data.get("content_type", "Book"),
            category=category,
            description=description,
            raw_text=text,
            audience=audience,
        )
        if ok:
            st.session_state.library_audience = audience
            st.session_state.page = audience
            st.rerun()
        else:
            st.warning(result)


# =========================================================
# READER
# =========================================================

def render_tts_player(text_to_read: str):
    """
    Browser-based speech player with touch seek, pause/resume,
    rewind/forward by text chunk, and multilingual female-voice preference.

    Web Speech API does not expose a true MP3-style timeline, so the seek bar
    maps to reading chunks across the chapter. Dragging the bar jumps to the
    closest chunk and continues from there when playing.
    """
    safe_text = json.dumps(text_to_read, ensure_ascii=False)

    player_html = r"""
    <!doctype html>
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        * {
            box-sizing: border-box;
        }

        html, body {
            margin: 0;
            padding: 0;
            background: transparent;
            font-family: Arial, sans-serif;
            color: #75655D;
        }

        .audio-shell {
            width: 100%;
            border-radius: 16px;
            padding: 7px 9px 6px;
            background:
                linear-gradient(
                    118deg,
                    #FBE1EA 0%,
                    #F4C8D8 28%,
                    #F8D7CE 56%,
                    #EEDAF1 78%,
                    #F7E5EE 100%
                );
            border: 1px solid #E6BBCB;
            box-shadow: 0 5px 14px rgba(232, 183, 200, .22);
        }

        .top-line {
            display: flex;
            align-items: center;
            gap: 8px;
            min-width: 0;
        }

        .round-btn {
            appearance: none;
            -webkit-appearance: none;
            width: 32px;
            min-width: 32px;
            height: 32px;
            border-radius: 50%;
            border: 1px solid #DDB4C4;
            background: rgba(255, 255, 255, .62);
            color: #75655D;
            font-size: 16px;
            font-weight: 700;
            line-height: 1;
            padding: 0;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            touch-action: manipulation;
            -webkit-tap-highlight-color: transparent;
        }

        .round-btn:hover,
        .round-btn:active,
        .round-btn:focus,
        .round-btn:focus-visible {
            background: rgba(255, 255, 255, .82);
            color: #75655D;
            border-color: #D6A9BB;
            outline: 2px solid rgba(220, 200, 161, .75);
            outline-offset: 1px;
        }

        .seek-wrap {
            flex: 1;
            min-width: 0;
            display: flex;
            align-items: center;
            gap: 7px;
        }

        .seek {
            width: 100%;
            min-width: 60px;
            height: 24px;
            background: transparent;
            accent-color: #DFA8BD;
            cursor: pointer;
            touch-action: pan-x;
        }

        .seek::-webkit-slider-runnable-track {
            height: 7px;
            border-radius: 999px;
            background: rgba(255,255,255,.78);
            border: 1px solid #E1BCCB;
        }

        .seek::-webkit-slider-thumb {
            -webkit-appearance: none;
            width: 20px;
            height: 20px;
            border-radius: 50%;
            margin-top: -7px;
            background: #DFA8BD;
            border: 2px solid #FFF9FB;
            box-shadow: 0 2px 5px rgba(117, 101, 93, .16);
        }

        .seek::-moz-range-track {
            height: 7px;
            border-radius: 999px;
            background: rgba(255,255,255,.78);
            border: 1px solid #E1BCCB;
        }

        .seek::-moz-range-thumb {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            background: #DFA8BD;
            border: 2px solid #FFF9FB;
        }

        .percent {
            min-width: 34px;
            text-align: right;
            font-size: 12px;
            font-weight: 700;
            color: #8A7468;
        }

        .hint {
            margin-top: 2px;
            padding-left: 2px;
            font-size: 10px;
            color: #8A7468;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        @media (max-width: 430px) {
            .audio-shell {
                padding: 7px 7px 6px;
                border-radius: 15px;
            }

            .top-line {
                gap: 6px;
            }

            .round-btn {
                width: 31px;
                min-width: 31px;
                height: 31px;
                font-size: 14px;
            }

            .percent {
                min-width: 31px;
                font-size: 11px;
            }
        }
    </style>
    </head>

    <body>
        <div class="audio-shell">
            <div class="top-line">
                <button id="rewind" class="round-btn" aria-label="Rewind">↶</button>
                <button id="playPause" class="round-btn" aria-label="Play or pause">▶</button>
                <button id="forward" class="round-btn" aria-label="Forward">↷</button>

                <div class="seek-wrap">
                    <input
                        id="seek"
                        class="seek"
                        type="range"
                        min="0"
                        max="100"
                        step="1"
                        value="0"
                        aria-label="Reading position"
                    />
                    <span id="percent" class="percent">0%</span>
                </div>
            </div>
            <div id="hint" class="hint">Read aloud · drag the bar to move through the chapter</div>
        </div>

        <script>
            const fullText = __TEXT_JSON__;
            const synth = window.speechSynthesis;

            const playPauseButton = document.getElementById("playPause");
            const rewindButton = document.getElementById("rewind");
            const forwardButton = document.getElementById("forward");
            const seek = document.getElementById("seek");
            const percent = document.getElementById("percent");
            const hint = document.getElementById("hint");

            let chunks = [];
            let voices = [];
            let currentIndex = 0;
            let isPlaying = false;
            let isPaused = false;
            let runId = 0;

            function langFor(text) {
                if (/[\u0E00-\u0E7F]/.test(text)) return "th-TH";
                if (/[\u4E00-\u9FFF]/.test(text)) return "zh-CN";
                return "en-US";
            }

            function splitByScript(text) {
                const lines = text
                    .split(/\n+/)
                    .map(s => s.trim())
                    .filter(Boolean);

                const result = [];

                for (const line of lines) {
                    const parts = line.match(
                        /[\u0E00-\u0E7F][\u0E00-\u0E7F0-9\s.,!?;:()\-–—/+&%]+|[\u4E00-\u9FFF][\u4E00-\u9FFF0-9\s，。！？；：、（）\-–—/+&%]+|[A-Za-z][A-Za-z0-9\s.,!?;:'"()\-–—/+&%]*/g
                    );

                    const usable = (parts && parts.length) ? parts : [line];

                    for (const rawPart of usable) {
                        const part = rawPart.trim();
                        if (!part) continue;

                        // Keep individual speech chunks reasonably short for mobile browsers.
                        if (part.length <= 190) {
                            result.push(part);
                        } else {
                            for (let start = 0; start < part.length; start += 180) {
                                const piece = part.slice(start, start + 180).trim();
                                if (piece) result.push(piece);
                            }
                        }
                    }
                }

                return result.length ? result : [text];
            }

            function femaleVoiceHints(lang) {
                const code = lang.toLowerCase();

                if (code.startsWith("th")) {
                    return ["premwadee", "kanya", "female", "woman"];
                }

                if (code.startsWith("zh")) {
                    return [
                        "xiaoxiao", "xiaoyi", "xiaohan", "xiaomeng", "xiaomo",
                        "xiaorui", "xiaoshuang", "xiaoxuan", "huihui", "yaoyao",
                        "meijia", "female", "woman"
                    ];
                }

                return [
                    "jenny", "aria", "zira", "samantha", "victoria", "karen",
                    "susan", "hazel", "ava", "emma", "female", "woman"
                ];
            }

            function knownMaleHints(lang) {
                const code = lang.toLowerCase();

                if (code.startsWith("th")) {
                    return ["niwat", "pattara", "male", "man"];
                }

                if (code.startsWith("zh")) {
                    return ["yunxi", "yunyang", "yunjian", "yunhao", "kangkang", "male", "man"];
                }

                return ["david", "mark", "guy", "george", "daniel", "james", "male", "man"];
            }

            function pickVoice(lang) {
                const exactLang = lang.toLowerCase();
                const prefix = exactLang.slice(0, 2);

                const localeVoices = voices.filter(v => {
                    const voiceLang = (v.lang || "").toLowerCase();
                    return voiceLang === exactLang || voiceLang.startsWith(prefix);
                });

                const femaleHints = femaleVoiceHints(lang);
                const maleHints = knownMaleHints(lang);

                for (const voiceHint of femaleHints) {
                    const match = localeVoices.find(v =>
                        (v.name || "").toLowerCase().includes(voiceHint)
                    );
                    if (match) return match;
                }

                return localeVoices.find(v => {
                    const name = (v.name || "").toLowerCase();
                    return !maleHints.some(voiceHint => name.includes(voiceHint));
                }) || null;
            }

            function loadVoices() {
                voices = synth.getVoices();

                if (!voices.length) {
                    synth.onvoiceschanged = () => {
                        voices = synth.getVoices();
                    };
                }
            }

            function clampIndex(index) {
                return Math.max(0, Math.min(index, Math.max(chunks.length - 1, 0)));
            }

            function percentageForIndex(index) {
                if (chunks.length <= 1) return index > 0 ? 100 : 0;
                return Math.round((index / (chunks.length - 1)) * 100);
            }

            function indexForPercentage(value) {
                if (chunks.length <= 1) return 0;
                return clampIndex(
                    Math.round((Number(value) / 100) * (chunks.length - 1))
                );
            }

            function updateUI() {
                const p = percentageForIndex(currentIndex);
                seek.value = String(p);
                percent.textContent = p + "%";
                playPauseButton.textContent = isPlaying ? "⏸" : "▶";

                if (isPlaying) {
                    hint.textContent = "Reading · drag the bar anytime to jump";
                } else if (isPaused) {
                    hint.textContent = "Paused · press play to continue";
                } else {
                    hint.textContent = "Read aloud · drag the bar to move through the chapter";
                }
            }

            function cancelSpeech() {
                runId += 1;
                synth.cancel();
                isPlaying = false;
                isPaused = false;
            }

            function speakChunk(localRunId) {
                if (!isPlaying || localRunId !== runId || !chunks.length) return;

                currentIndex = clampIndex(currentIndex);
                updateUI();

                const chunk = chunks[currentIndex];
                const lang = langFor(chunk);
                const utterance = new SpeechSynthesisUtterance(chunk);

                utterance.lang = lang;
                utterance.rate = lang === "th-TH" ? 0.88 : 0.92;
                utterance.pitch = 1.0;

                const voice = pickVoice(lang);
                if (voice) utterance.voice = voice;

                utterance.onend = () => {
                    if (!isPlaying || localRunId !== runId) return;

                    if (currentIndex < chunks.length - 1) {
                        currentIndex += 1;
                        speakChunk(localRunId);
                    } else {
                        isPlaying = false;
                        isPaused = false;
                        currentIndex = Math.max(chunks.length - 1, 0);
                        seek.value = "100";
                        percent.textContent = "100%";
                        playPauseButton.textContent = "▶";
                        hint.textContent = "Finished · drag back to replay any section";
                    }
                };

                utterance.onerror = () => {
                    if (!isPlaying || localRunId !== runId) return;

                    if (currentIndex < chunks.length - 1) {
                        currentIndex += 1;
                        speakChunk(localRunId);
                    } else {
                        isPlaying = false;
                        updateUI();
                    }
                };

                synth.speak(utterance);
            }

            function startFromCurrent() {
                if (!("speechSynthesis" in window) || !chunks.length) {
                    hint.textContent = "Speech is not available in this browser.";
                    return;
                }

                synth.cancel();
                runId += 1;
                const localRunId = runId;

                isPlaying = true;
                isPaused = false;
                updateUI();

                // Small timeout avoids mobile browsers racing cancel() and speak().
                setTimeout(() => speakChunk(localRunId), 35);
            }

            function togglePlayPause() {
                if (isPlaying) {
                    synth.pause();
                    isPlaying = false;
                    isPaused = true;
                    updateUI();
                    return;
                }

                if (isPaused && synth.paused) {
                    synth.resume();
                    isPlaying = true;
                    isPaused = false;
                    updateUI();
                    return;
                }

                startFromCurrent();
            }

            function jumpBy(delta) {
                const wasPlaying = isPlaying;

                synth.cancel();
                runId += 1;
                isPlaying = false;
                isPaused = false;

                currentIndex = clampIndex(currentIndex + delta);
                updateUI();

                if (wasPlaying) {
                    startFromCurrent();
                }
            }

            function seekTo(value) {
                const wasPlaying = isPlaying;

                synth.cancel();
                runId += 1;
                isPlaying = false;
                isPaused = false;

                currentIndex = indexForPercentage(value);
                updateUI();

                if (wasPlaying) {
                    startFromCurrent();
                }
            }

            chunks = splitByScript(fullText);
            loadVoices();
            updateUI();

            playPauseButton.addEventListener("click", togglePlayPause);
            rewindButton.addEventListener("click", () => jumpBy(-1));
            forwardButton.addEventListener("click", () => jumpBy(1));

            seek.addEventListener("input", e => {
                const value = Number(e.target.value);
                percent.textContent = Math.round(value) + "%";
            });

            seek.addEventListener("change", e => {
                seekTo(e.target.value);
            });

            // Cancel speech if Streamlit removes/replaces this player iframe.
            window.addEventListener("beforeunload", () => {
                synth.cancel();
            });
        </script>
    </body>
    </html>
    """

    player_html = player_html.replace("__TEXT_JSON__", safe_text)
    components.html(player_html, height=68, scrolling=False)


PASTEL_VOCAB_GRADIENTS = [
    # Richer pastel tones: still soft, but dark enough to read comfortably on mobile.
    "linear-gradient(90deg,#C57F9A 0%,#A98CC7 55%,#7FAFC9 100%)",
    "linear-gradient(90deg,#89A77D 0%,#B79363 52%,#C7849D 100%)",
    "linear-gradient(90deg,#7FAFC8 0%,#9C85BC 50%,#C47E98 100%)",
    "linear-gradient(90deg,#B88B58 0%,#C47792 52%,#9B84BA 100%)",
    "linear-gradient(90deg,#86A57C 0%,#7EA9BD 48%,#9A84B8 100%)",
    "linear-gradient(90deg,#C5829C 0%,#B99463 48%,#88A57C 100%)",
    "linear-gradient(90deg,#9B83BB 0%,#C67D98 48%,#79A9C2 100%)",
    "linear-gradient(90deg,#789F98 0%,#B49161 50%,#9981B5 100%)",
]


def _chapter_nav_label(title: str) -> str:
    # Keep navigator pills compact while preserving recognizable chapter names.
    cleaned = re.sub(r"\s+", " ", str(title or "Chapter")).strip()
    cleaned = re.sub(r"^(Day\s+\d+)\s*[—–-]\s*", r"\1 · ", cleaned, flags=re.I)
    if len(cleaned) > 34:
        cleaned = cleaned[:33].rstrip() + "…"
    return cleaned


def render_chapter_navigator(book, current_index: int):
    """Render stable HTML pills that look identical on local, Cloud, and mobile.

    We intentionally avoid st.radio here because Streamlit Cloud and local builds
    can produce different radio DOM structures, which made the mobile labels wrap
    into fragments. Links use query parameters and are handled in main().
    """
    chapters = book.get("chapters", [])
    if len(chapters) <= 1:
        return

    pills = []
    for i, chapter in enumerate(chapters):
        label = escape(_chapter_nav_label(chapter.get("chapter_title", f"Chapter {i + 1}")))
        active = " active" if i == current_index else ""
        href = f"?reader_book={book['book_id']}&reader_chapter={i}"
        pills.append(
            f'<a class="chapter-nav-pill{active}" href="{href}" target="_self">'
            f'<span class="chapter-nav-dot"></span><span>{label}</span></a>'
        )

    st.markdown(
        '<div class="chapter-nav-scroll">' + ''.join(pills) + '</div>',
        unsafe_allow_html=True,
    )


def _extract_thai_gloss(roots: str) -> str:
    # Pull a concise Thai meaning from the existing bilingual root note when possible.
    roots = str(roots or "").strip()
    if not roots:
        return ""

    if "→" in roots:
        tail = roots.rsplit("→", 1)[-1].strip()
    elif " / " in roots:
        tail = roots.split(" / ", 1)[-1].strip()
    else:
        tail = roots

    tail = tail.split(";", 1)[0].strip()
    if "/" in tail:
        candidate = tail.rsplit("/", 1)[-1].strip()
        if re.search(r"[\u0E00-\u0E7F]", candidate):
            return candidate.strip(" .")

    match = re.search(r"[\u0E00-\u0E7F][\u0E00-\u0E7F\s,()\-–—]*", tail)
    return match.group(0).strip(" ,.-") if match else ""


def _parse_vocab_chapter(content: str):
    # Parse the Chinese book's existing English / Chinese / Pinyin / Roots structure.
    lines = str(content or "").splitlines()
    positions = []

    for line_index, line in enumerate(lines):
        match = re.match(r"^\s*(\d+)\.\s+(.+?)\s*$", line)
        if match:
            positions.append((line_index, int(match.group(1)), match.group(2).strip()))

    if not positions:
        return [], []

    first_entry_line = positions[0][0]
    preface = [line.strip() for line in lines[:first_entry_line] if line.strip()]
    if preface:
        preface = preface[1:]

    entries = []
    for position_index, (line_index, number, english) in enumerate(positions):
        end_line = (
            positions[position_index + 1][0]
            if position_index + 1 < len(positions)
            else len(lines)
        )
        segment = [line.strip() for line in lines[line_index + 1:end_line] if line.strip()]
        if len(segment) < 2:
            continue

        chinese = segment[0]
        pinyin = segment[1]
        roots = " ".join(segment[2:]).strip()
        entries.append(
            {
                "number": number,
                "english": english,
                "chinese": chinese,
                "pinyin": pinyin,
                "roots": roots,
                "thai": _extract_thai_gloss(roots),
            }
        )

    return preface, entries


def _sentence_case_label(text: str) -> str:
    text = str(text or "").strip()
    if not text:
        return text
    if text[0].islower():
        return text[0].upper() + text[1:]
    return text


def render_vocab_chapter(chapter, font_size: int):
    # Compact Layout B: large Chinese, Pinyin, meaning, then root explanation.
    preface, entries = _parse_vocab_chapter(chapter.get("content", ""))

    if not entries:
        st.markdown(
            f'''<div class="reader-paper"><div class="reader-content" style="font-size:{font_size}px;">{escape(chapter.get("content", ""))}</div></div>''',
            unsafe_allow_html=True,
        )
        return

    intro_html = ""
    if preface:
        intro_text = " ".join(preface)
        intro_html = f'<div class="vocab-intro">{escape(intro_text)}</div>'

    cards = []
    for item in entries:
        gradient_index = (
            sum(ord(ch) for ch in item["chinese"]) + item["number"] * 7
        ) % len(PASTEL_VOCAB_GRADIENTS)
        gradient = PASTEL_VOCAB_GRADIENTS[gradient_index]
        english = _sentence_case_label(item["english"])
        thai = item["thai"]
        meaning = f"{english} · {thai}" if thai else english
        roots = item["roots"] or "Roots: —"

        # Keep every card as one continuous HTML fragment. Streamlit Markdown can
        # interpret indented/newline-separated HTML between sibling cards as a
        # Markdown code block; minifying the fragment avoids that parser edge case.
        cards.append(
            f'<div class="vocab-card">'
            f'<div class="vocab-topline">'
            f'<div class="vocab-number">{item["number"]}</div>'
            f'<div class="vocab-wordblock">'
            f'<div class="vocab-chinese" style="--vocab-gradient:{gradient};">{escape(item["chinese"])}</div>'
            f'<div class="vocab-pinyin">{escape(item["pinyin"])}</div>'
            f'<div class="vocab-meaning">{escape(meaning)}</div>'
            f'</div>'
            f'</div>'
            f'<div class="vocab-roots">{escape(roots)}</div>'
            f'</div>'
        )

    # Keep the entire vocabulary reader on one HTML line. No blank lines or
    # indentation are allowed between cards, otherwise Markdown may render the
    # later cards literally as <div> source code.
    vocab_html = (
        f'<div class="vocab-reader" style="--reader-font-size:{font_size}px;">'
        f'{intro_html}'
        f'<div class="vocab-list">{"".join(cards)}</div>'
        f'</div>'
    ).replace("\n", "")
    st.markdown(vocab_html, unsafe_allow_html=True)


def render_reader():
    book = current_book()
    if not book or not book.get("chapters"):
        st.warning("This book does not contain readable chapters.")
        return

    idx = st.session_state.current_chapter_index
    idx = max(0, min(idx, len(book["chapters"]) - 1))
    st.session_state.current_chapter_index = idx
    st.session_state.last_read[book["book_id"]] = idx

    chapter = book["chapters"][idx]
    progress = calculate_progress(book, idx)

    # Main reader controls stay at the top so they are reachable without scrolling.
    # Chapter movement is handled by the sub-category navigator beneath Read Aloud.
    with st.container(key="reader_top_actions"):
        with st.container(key="reader_fixed_audio"):
            with st.container(key="reader_audio_control_row"):
                audio_cols = st.columns([1, 1, 1, 9], gap="small")

                with audio_cols[0]:
                    if st.button("A−", key="font_minus"):
                        st.session_state.font_size = max(15, st.session_state.font_size - 2)
                        st.rerun()

                with audio_cols[1]:
                    if st.button("A+", key="font_plus"):
                        st.session_state.font_size = min(27, st.session_state.font_size + 2)
                        st.rerun()

                with audio_cols[2]:
                    bookmarked = is_bookmarked(book["book_id"], chapter["chapter_id"])
                    if st.button("♥" if bookmarked else "♡", key="bookmark_current"):
                        toggle_bookmark(book, chapter)
                        st.rerun()

                with audio_cols[3]:
                    render_tts_player(chapter["content"])

            # Sub-category / chapter pills stay directly beneath Read Aloud.
            render_chapter_navigator(book, idx)


    st.markdown(f"## {escape(chapter['chapter_title'])}")

    render_progress(progress)
    st.caption(f"{progress}% complete · {book['title']}")

    with st.container(key="reader_bunny_art"):
        bunny_cols = st.columns([1, 1.65, 1], gap="small")
        with bunny_cols[1]:
            st.image(
                str(ASSET_DIR / "reader_bunny.jpg"),
                use_container_width=True,
            )

    if st.session_state.reader_notice:
        st.markdown(
            f'<div class="pastel-note pastel-success">{escape(st.session_state.reader_notice)}</div>',
            unsafe_allow_html=True,
        )
        st.session_state.reader_notice = ""

    font_size = st.session_state.font_size
    if book.get("book_id") == "chinese-for-technical-product-managers":
        render_vocab_chapter(chapter, font_size)
    else:
        reader_html = (
            f'<div class="reader-paper">'
            f'<div class="reader-content" style="font-size:{font_size}px;">'
            f'{escape(chapter["content"])}'
            f'</div></div>'
        )
        st.markdown(reader_html, unsafe_allow_html=True)




# =========================================================
# BOOKMARKS
# =========================================================
def render_bookmarks():
    st.markdown("## Bookmarks")

    if not st.session_state.bookmarks:
        st.markdown(
            """
            <div class="empty-state">
                <div style="font-size:48px;">🐰♡</div>
                <div class="book-title">No bookmarks yet.</div>
                <div class="muted">Save a chapter while reading and it will appear here.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    for idx, item in enumerate(st.session_state.bookmarks):
        st.markdown(
            f"""
            <div class="soft-card">
                <div class="book-title">{escape(item["book_title"])}</div>
                <div class="muted">{escape(item["chapter_title"])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Open Bookmark", key=f"open_bm_{idx}"):
            open_book(item["book_id"], item["chapter_index"])
            st.rerun()


# =========================================================
# READER DEEP-LINK / CHAPTER NAVIGATION
# =========================================================
def apply_reader_query_navigation():
    """Apply chapter-pill links before rendering the page."""
    try:
        book_id = st.query_params.get("reader_book")
        chapter_raw = st.query_params.get("reader_chapter")
    except Exception:
        return

    if not book_id or chapter_raw is None:
        return

    book = get_book(str(book_id))
    if not book or not book.get("chapters"):
        return

    try:
        chapter_index = int(chapter_raw)
    except (TypeError, ValueError):
        return

    chapter_index = max(0, min(chapter_index, len(book["chapters"]) - 1))
    st.session_state.current_book_id = book["book_id"]
    st.session_state.library_audience = book_audience(book)
    st.session_state.page = "Reader"

    if st.session_state.current_chapter_index != chapter_index:
        st.session_state.current_chapter_index = chapter_index
        st.session_state.last_read[book["book_id"]] = chapter_index
        ok, error = persist_progress(book["book_id"], chapter_index)
        if not ok and storage_is_configured():
            st.session_state.storage_error = error or "Could not save reading progress."


# =========================================================
# APP ROUTER
# =========================================================
def main():
    initialize_state()
    apply_reader_query_navigation()
    inject_css()

    render_navigation()

    page = st.session_state.page
    if page == "Home":
        render_home()
    elif page == "Kids":
        render_library("Kids")
    elif page == "Mommy":
        render_library("Mommy")
    elif page == "Library":
        # Backward compatibility for an older browser session.
        st.session_state.page = st.session_state.get("library_audience", "Mommy")
        render_library(st.session_state.page)
    elif page == "Add":
        render_add_content()
    elif page == "Bookmarks":
        render_bookmarks()
    elif page == "Reader":
        render_reader()
    else:
        st.session_state.page = "Home"
        render_home()

    render_storybook_footer()


if __name__ == "__main__":
    main()

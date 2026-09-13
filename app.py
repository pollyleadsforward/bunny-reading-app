
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
DEMO_BOOKS = [{'book_id': 'product-development-interview-prep',
  'title': 'Product Development Interview Prep',
  'subtitle': 'Head of Personal Loan Product Management · Full Review',
  'author': 'Personal Study Library',
  'content_type': 'Book',
  'category': 'Interview Prep',
  'description': 'JD-based interview review covering product strategy, customer needs, production '
                 'quality, regulation, market, PMF, risk, incidents, and head-level questions.',
  'cover_emoji': '📘',
  'chapters': [{'chapter_id': 'tell-me-about-yourself',
                'chapter_title': 'Tell Me About Yourself',
                'order': 1,
                'content': 'TELL ME ABOUT YOURSELF\n'
                           '\n'
                           '中文\n'
                           '\n'
                           '目前，我负责数字贷款产品，主要关注从需求、测试、上线到上线后问题管理的端到端产品流程。\n'
                           '\n'
                           '\n'
                           '我的经验主要有三个方面。\n'
                           '\n'
                           '\n'
                           '第一是产品管理。我和不同团队合作，不断改善客户流程和产品体验。\n'
                           '\n'
                           '\n'
                           '第二是生产环境和系统事故管理。发生问题时，我会先了解客户和业务影响，再分析原因，并推动临时方案和长期解决方案。\n'
                           '\n'
                           '\n'
                           '第三是风险和合规。我会和风险、合规、法务、运营和技术团队合作，确保产品满足客户需求，也符合相关规定。\n'
                           '\n'
                           '\n'
                           '我对这个职位很感兴趣，因为我希望从数字贷款进一步扩展到更全面的个人贷款产品管理，包括产品策略、客户需求、市场机会和产品市场匹配。\n'
                           '\n'
                           '\n'
                           '我的优势是能够从端到端看产品，把客户、业务、风险和技术连接起来，并把生产环境中的问题转化为持续改善产品的机会。\n'
                           '\n'
                           '\n'
                           'ENGLISH\n'
                           '\n'
                           'Currently, I work in Digital Lending, where I look after the '
                           'end-to-end product journey, from requirements and testing through '
                           'production and post-launch issue management.\n'
                           '\n'
                           'My experience is mainly in three areas. First, product management, '
                           'working across functions to develop and continuously improve the '
                           'customer journey. Second, production and incident management, where I '
                           'focus on understanding customer and business impact, identifying root '
                           'causes, and driving both workarounds and long-term solutions. Third, '
                           'risk and compliance, where I work closely with Risk, Compliance, '
                           'Legal, Operations, and Technology to ensure that the product serves '
                           'customer needs while operating within the right controls and '
                           'regulations.\n'
                           '\n'
                           'What interests me about this role is the opportunity to expand from '
                           'Digital Lending into a broader Personal Loan Product Management scope, '
                           'including product strategy, customer needs, portfolio performance, '
                           'market opportunities, and product-market fit.\n'
                           '\n'
                           'I believe my key strength is my ability to look at a product end to '
                           'end, connect customer, business, risk, and technology perspectives, '
                           'and turn production issues into opportunities for continuous product '
                           'improvement.\n'
                           '\n'
                           '\n'
                           'ภาษาไทย\n'
                           '\n'
                           'ปัจจุบันดิฉันดูแลงานด้าน Digital Lending โดยดูแล end-to-end product '
                           'journey ตั้งแต่การพัฒนา requirement การทดสอบ ไปจนถึง production '
                           'และการดูแลปัญหาหลังจากระบบขึ้นใช้งานค่ะ\n'
                           '\n'
                           'ประสบการณ์หลักของดิฉันอยู่ใน 3 ด้านค่ะ หนึ่ง คือ Product Management '
                           'การทำงานร่วมกับหลายทีมเพื่อพัฒนาและปรับปรุง customer journey สอง คือ '
                           'Production & Incident Management โดยเฉพาะการวิเคราะห์ผลกระทบ หา root '
                           'cause และวางทั้ง workaround และ long-term solution และสาม คือ Risk & '
                           'Compliance ซึ่งต้องทำงานร่วมกับ Risk, Compliance, Legal, Operations '
                           'และ Technology เพื่อให้ product '
                           'ตอบโจทย์ลูกค้าและอยู่ภายใต้ข้อกำหนดที่เหมาะสม\n'
                           '\n'
                           'สิ่งที่ดิฉันสนใจสำหรับตำแหน่งนี้ คือโอกาสที่จะขยายจากการดูแล Digital '
                           'Lending ไปสู่การดูแล Personal Loan Product ในภาพที่กว้างขึ้น ทั้งด้าน '
                           'strategy, customer needs, portfolio performance, market opportunity '
                           'และ product-market fit\n'
                           '\n'
                           'ดิฉันคิดว่าจุดแข็งของตัวเองคือการมองปัญหาแบบ end to end เชื่อม '
                           'customer, business, risk และ technology เข้าด้วยกัน '
                           'และเปลี่ยนปัญหาที่เกิดขึ้นใน production ให้กลายเป็นโอกาสในการพัฒนา '
                           'product ให้ดีขึ้นค่ะ'},
               {'chapter_id': 'product-strategy',
                'chapter_title': 'Product Strategy',
                'order': 2,
                'content': 'PRODUCT STRATEGY & MANAGEMENT\n'
                           '\n'
                           '1) ถ้าให้คุณวาง Strategy สำหรับ Personal Loan คุณจะเริ่มจากอะไร?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันจะเริ่มจาก 5 เรื่องค่ะ: Target Customer, Customer Need, Risk, Value '
                           'Proposition และ Economics แล้วจึงออกแบบ product, pricing, channel และ '
                           'process ให้สอดคล้องกัน ไม่เริ่มจากคำถามว่า “เราจะขายอะไร” แต่เริ่มจาก '
                           '“ลูกค้ากลุ่มไหนมี unmet need ที่เราสามารถให้บริการได้อย่าง profitable '
                           'และ sustainable”\n'
                           '\n'
                           '\n'
                           '4) Personal Loan ที่ดีควรแข่งขันด้วยอะไร นอกจากดอกเบี้ย?\n'
                           '\n'
                           'Answer:\n'
                           'แข่งขันได้หลายเรื่องค่ะ เช่น approval speed, ease of application, '
                           'transparency, flexible repayment, relevant credit limit และ customer '
                           'experience ถ้าแข่งขันด้วยราคาอย่างเดียวจะเข้าสู่ price war ได้ง่าย\n'
                           '\n'
                           '\n'
                           '7) Product Strategy กับ Credit Policy ต่างกันอย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'Product Strategy ตอบว่าเราจะให้ใคร อะไร ด้วย value proposition แบบไหน '
                           'และทำไมถึงน่าสนใจเชิงธุรกิจ ส่วน Credit Policy กำหนดว่า risk '
                           'แบบไหนที่ธนาคารยอมรับ และเงื่อนไขอนุมัติเป็นอย่างไร '
                           'สองเรื่องต้องออกแบบร่วมกัน\n'
                           '\n'
                           '\n'
                           '8) คุณตั้ง KPI ของ Personal Loan อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันจะไม่ดูแค่ยอดขาย แต่ดูครบทั้ง Growth + Customer + Risk + Economics '
                           '+ Operations เช่น application, approval, booking/disbursement, '
                           'conversion, revenue, risk-adjusted return, delinquency/NPL, complaint '
                           'และ SLA\n'
                           '\n'
                           '\n'
                           '9) ถ้ายอด Loan Booking โตมาก คุณถือว่าประสบความสำเร็จหรือยัง?\n'
                           '\n'
                           'Answer:\n'
                           'ยังค่ะ Growth อย่างเดียวไม่พอ ต้องดู portfolio quality และ '
                           'profitability ด้วย ถ้า booking โตแต่ NPL, complaint หรือ acquisition '
                           'cost โตเร็วกว่า แปลว่า growth นั้นอาจไม่ sustainable'},
               {'chapter_id': 'customer-needs-discovery',
                'chapter_title': 'Customer Needs & Discovery',
                'order': 3,
                'content': 'CUSTOMER NEEDS & DISCOVERY\n'
                           '\n'
                           '2) คุณจะรู้ได้อย่างไรว่าลูกค้าต้องการอะไร?\n'
                           '\n'
                           'Answer:\n'
                           'ดูทั้ง customer data และ customer voice ค่ะ เช่น application funnel, '
                           'drop-off, approval/rejection reason, complaints, usage behavior รวมถึง '
                           'qualitative research แล้วเอาข้อมูลมาหา pain point ก่อนพัฒนา product\n'
                           '\n'
                           '\n'
                           '3) ถ้ามีหลาย Customer Segment จะเลือกกลุ่มไหนก่อน?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันจะดู 4 มิติ ได้แก่ Market attractiveness, Customer need, Risk '
                           'profile และ Bank capability กลุ่มที่มี need ชัด มี market size เพียงพอ '
                           'risk สามารถบริหารได้ และธนาคารมี data/capability รองรับ '
                           'จะเป็นกลุ่มที่ควร prioritize ก่อน\n'
                           '\n'
                           '\n'
                           '17) คุณจะติดตามการเปลี่ยนแปลงของ Customer Behavior อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ดูทั้ง application behavior, channel usage, transaction pattern, '
                           'repayment behavior, complaints และ research '
                           'แล้วดูว่าการเปลี่ยนแปลงนั้นมีผลต่อ product proposition หรือ risk '
                           'อย่างไร\n'
                           '\n'
                           '\n'
                           '18) ถ้าพฤติกรรมลูกค้าเปลี่ยน คุณจะรู้ได้อย่างไรว่าต้องแก้ Product?\n'
                           '\n'
                           'Answer:\n'
                           'ต้องดูว่า change นั้นเป็น temporary noise หรือ structural change '
                           'ถ้าเกิดต่อเนื่องและกระทบ conversion, profitability, risk หรือ customer '
                           'satisfaction จึงควรพิจารณา product redesign'},
               {'chapter_id': 'product-development',
                'chapter_title': 'Product Development',
                'order': 4,
                'content': 'PRODUCT DEVELOPMENT & CONTINUOUS IMPROVEMENT\n'
                           '\n'
                           '10) Continuous Improvement ทำอย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันจะสร้าง feedback loop จาก production data: identify pain point → '
                           'hypothesis → improvement → test → measure และทำต่อเนื่อง ไม่รอให้มี '
                           'major project ถึงจะปรับ product\n'
                           '\n'
                           '\n'
                           '24) เวลา Analyze Competitor คุณดูอะไร?\n'
                           '\n'
                           'Answer:\n'
                           'ไม่ดูแค่ interest rate ค่ะ ฉันจะดู target segment, eligibility, credit '
                           'limit, pricing, tenor, application journey, approval speed, channel, '
                           'campaign และ value proposition\n'
                           '\n'
                           '\n'
                           '25) ถ้าคู่แข่งลดดอกเบี้ยแรง คุณจะลดตามไหม?\n'
                           '\n'
                           'Answer:\n'
                           'ไม่จำเป็นค่ะ ต้องรู้ก่อนว่าเขากำลัง target ลูกค้ากลุ่มไหน และ '
                           'economics ของเรารองรับหรือไม่ บางครั้งการตอบด้วย better experience, '
                           'targeted pricing หรือ differentiated segment มีเหตุผลกว่าการลดราคาทั้ง '
                           'portfolio\n'
                           '\n'
                           '\n'
                           '26) คุณหา Opportunity จาก Market Trend อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'มองหาจุดที่มี customer need เพิ่มขึ้น แต่ existing solution '
                           'ยังตอบไม่ดี แล้วประเมินร่วมกับ bank capability, data advantage และ '
                           'risk appetite\n'
                           '\n'
                           '\n'
                           '28) Competitor ทำ Feature ใหม่ เราควรทำตามไหม?\n'
                           '\n'
                           'Answer:\n'
                           'ไม่ควร copy โดยอัตโนมัติ ต้องถามว่า customer problem คืออะไร และ '
                           'feature นั้นสร้าง value จริงหรือไม่ ถ้าเหมาะกับลูกค้าและ strategy '
                           'ของเรา จึงค่อยนำมาปรับใช้'},
               {'chapter_id': 'testing-uat-deployment',
                'chapter_title': 'Testing / UAT / Deployment',
                'order': 5,
                'content': 'PRODUCT TESTING / UAT / PRODUCTION READINESS\n'
                           '\n'
                           '11) “Ensuring quality in production” ใน JD คุณตีความว่าอะไร?\n'
                           '\n'
                           'Answer:\n'
                           'ไม่ใช่แค่ระบบไม่ล่มค่ะ '
                           'แต่หมายถึงลูกค้าต้องได้รับผลลัพธ์ที่ถูกต้องตั้งแต่ application → '
                           'decision → documentation → disbursement → servicing รวมถึง data '
                           'accuracy, SLA, compliance และ customer communication\n'
                           '\n'
                           '\n'
                           '12) หลัง Go-live คุณดูอะไรเป็นอันดับแรก?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันจะดู critical funnel เช่น application success, approval result, '
                           'disbursement success, error rate, exception, customer complaint และ '
                           'reconciliation เทียบกับ baseline และ expected result\n'
                           '\n'
                           '\n'
                           '31) ก่อน Launch Product ใหม่ คุณจะ Test อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'เริ่มจาก hypothesis และ success criteria แล้วทำ pilot กับ segment '
                           'จำกัด กำหนด guardrail ด้าน risk/operations/customer impact ก่อนขยาย '
                           'scale\n'
                           '\n'
                           '\n'
                           '32) Product Test ต่างจาก UAT อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'UAT ตอบว่า “ระบบทำงานตาม requirement หรือไม่” แต่ Product Test ตอบว่า '
                           '“product นี้ตอบโจทย์ลูกค้าและธุรกิจหรือไม่” ระบบอาจผ่าน UAT แต่ '
                           'product ยังไม่ผ่าน PMF ก็ได้\n'
                           '\n'
                           '\n'
                           '34) คุณกำหนด Go / No-Go อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ดู critical defects, customer impact, regulatory issue, operational '
                           'readiness, workaround และ remaining risk ถ้ามี issue แต่มี workaround '
                           'และ controlled risk อาจ Go ได้ แต่ material customer/regulatory risk '
                           'ต้องแก้ก่อน'},
               {'chapter_id': 'product-market-fit',
                'chapter_title': 'Product-Market Fit',
                'order': 6,
                'content': 'PRODUCT-MARKET FIT (PMF)\n'
                           '\n'
                           '29) PMF ใน Personal Loan คืออะไร?\n'
                           '\n'
                           'Answer:\n'
                           'สำหรับ Lending ฉันมองว่า PMF ไม่ใช่แค่มีคนสมัครเยอะ แต่คือ ลูกค้าเห็น '
                           'value, ใช้ product จริง และ portfolio สร้าง economics ที่ดีภายใต้ '
                           'acceptable risk ต้องมีทั้ง customer fit และ business fit\n'
                           '\n'
                           '\n'
                           '30) จะวัด PMF ด้วยอะไร?\n'
                           '\n'
                           'Answer:\n'
                           'เช่น application demand, conversion, approval-to-booking, '
                           'utilization/drawdown, repeat behavior, complaints, customer '
                           'satisfaction, acquisition cost, portfolio quality และ profitability\n'
                           '\n'
                           '\n'
                           '33) ถ้า Pilot Conversion สูงมาก แต่ NPL เริ่มสูง คุณจะ Scale ไหม?\n'
                           '\n'
                           'Answer:\n'
                           'ยังไม่ scale ค่ะ ต้องเข้าใจก่อนว่า growth มาจาก segment ไหน และ risk '
                           'สูงเพราะอะไร อาจต้องปรับ eligibility, limit, pricing หรือ underwriting '
                           'ก่อน'},
               {'chapter_id': 'lending-risk',
                'chapter_title': 'Lending & Risk',
                'order': 7,
                'content': 'LENDING & RISK\n'
                           '\n'
                           '5) คุณจะออกแบบ Pricing อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันมองเป็น risk-based pricing ค่ะ ต้อง balance customer affordability, '
                           'expected loss, cost, acquisition expense และ target return '
                           'ลูกค้าที่มีข้อมูลและความเสี่ยงต่างกันไม่จำเป็นต้องได้ราคาเดียวกัน\n'
                           '\n'
                           '\n'
                           '6) ถ้าจะทำสินเชื่อให้ Self-employed คุณจะออกแบบอย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันจะไม่มอง Self-employed เป็นกลุ่มเดียว แต่แบ่งตาม quality of '
                           'evidence เช่น มีทะเบียน/ข้อมูลธุรกิจชัด, มี transaction data เช่น '
                           'QR/EDC หรือมีเพียง financial behavior จากบัญชี จากนั้นกำหนด '
                           'eligibility, limit และ pricing ตามระดับความเชื่อมั่นของข้อมูล\n'
                           '\n'
                           '\n'
                           '27) แล้วหา Risk จาก Market Trend อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ดูทั้ง credit deterioration, aggressive competition, customer '
                           'indebtedness, fraud pattern, regulatory direction และ cost of '
                           'acquisition เพราะบาง market growth อาจดู attractive แต่มี risk '
                           'ซ่อนอยู่\n'
                           '\n'
                           '\n'
                           '35) ถ้า Growth Target กับ Risk Target ขัดกัน คุณจะเลือกอะไร?\n'
                           '\n'
                           'Answer:\n'
                           'หน้าที่ของ Product ไม่ใช่เลือก Growth หรือ Risk แต่คือหา risk-adjusted '
                           'growth ค่ะ Growth ที่ทำลาย portfolio quality ไม่ sustainable '
                           'ขณะเดียวกัน risk control ที่ conservative เกินไปก็ทำให้เสีย '
                           'opportunity ดังนั้นต้องปรับ segment, pricing, limit หรือ criteria ให้ '
                           'balance'},
               {'chapter_id': 'incident-operations',
                'chapter_title': 'Incident & Operations',
                'order': 8,
                'content': 'INCIDENT & OPERATIONS\n'
                           '\n'
                           '13) ถ้ามี Incident หลัง Go-live คุณจะทำอย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'อันดับแรกคือ protect customer และ contain impact จากนั้นระบุ scope, '
                           'workaround, owner และ SLA แล้วค่อยทำ root cause และ permanent fix '
                           'หลังเหตุการณ์ต้องมี preventive action เพื่อไม่ให้เกิดซ้ำ\n'
                           '\n'
                           '\n'
                           '14) คุณแยก Major Incident กับ Minor Issue อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ดูจาก customer impact, financial impact, regulatory risk, number of '
                           'customers และ business continuity ไม่ใช่ดูจากความยากของ technical '
                           'issue อย่างเดียว\n'
                           '\n'
                           '\n'
                           '15) ถ้า System Issue กระทบลูกค้าแค่ 1 คน ยังต้อง investigate ไหม?\n'
                           '\n'
                           'Answer:\n'
                           'ต้องค่ะ เพราะต้องตอบให้ได้ว่าเป็น isolated case หรือ systemic logic '
                           'issue ถ้าเป็น logic issue '
                           'แม้วันนี้เจอคนเดียวก็อาจกระทบลูกค้ารายอื่นในอนาคต\n'
                           '\n'
                           '\n'
                           '16) Product Owner ต้องรู้ Technical แค่ไหน?\n'
                           '\n'
                           'Answer:\n'
                           'ไม่จำเป็นต้องเขียนระบบเอง แต่ต้องเข้าใจ end-to-end flow, data, '
                           'integration, business rule และ failure point มากพอที่จะตั้งคำถามกับ '
                           'Technology และประเมิน customer/business impact ได้'},
               {'chapter_id': 'leadership-collaboration',
                'chapter_title': 'Leadership & Collaboration',
                'order': 9,
                'content': 'REGULATION, LEADERSHIP & CROSS-FUNCTIONAL COLLABORATION\n'
                           '\n'
                           '19) ถ้ามีกฎใหม่ออกมา คุณจัดการอย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันจะทำ impact assessment ก่อนว่า policy, process, system, '
                           'communication และ existing customer มีอะไรได้รับผลกระทบ '
                           'จากนั้นทำงานกับ Compliance/Legal/Risk/Operations/IT เพื่อแปลง '
                           'regulation เป็น business requirement และ implementation plan\n'
                           '\n'
                           '\n'
                           '20) Product กับ Compliance ควรทำงานกันแบบไหน?\n'
                           '\n'
                           'Answer:\n'
                           'Compliance ไม่ควรเข้ามาเฉพาะตอนท้ายค่ะ ฉันชอบ involve ตั้งแต่ช่วง '
                           'design เพื่อให้ requirement ถูกตั้งแต่ต้น ลด rework '
                           'และทำให้ทีมเข้าใจด้วยว่า regulation ต้องการควบคุม risk อะไร\n'
                           '\n'
                           '\n'
                           '21) ถ้า Business อยาก Launch แต่ Compliance ยังมี Concern '
                           'คุณจะทำอย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ฉันจะไม่มองว่าเป็น Business vs Compliance แต่จะ clarify specific risk '
                           'และ requirement ก่อน แล้วดูว่ามี mitigation หรือ scope adjustment '
                           'ที่ทำให้ launch ได้อย่าง compliant หรือไม่ ถ้ายังมี material risk '
                           'ก็ไม่ควรฝืน Go\n'
                           '\n'
                           '\n'
                           '22) ถ้ากฎทำให้ Conversion ลดลง คุณจะทำอย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'Compliance requirement เป็น constraint ที่ต้องรักษา แต่เรายัง optimize '
                           'customer journey, wording, data collection และ process ได้ '
                           'เป้าหมายคือรักษากฎโดยลด unnecessary friction\n'
                           '\n'
                           '\n'
                           '23) คุณจะป้องกัน Product Requirement ผิดจาก Regulation ได้อย่างไร?\n'
                           '\n'
                           'Answer:\n'
                           'ต้องมี traceability ตั้งแต่ regulation → policy interpretation → '
                           'requirement → test case → production control เพื่อให้ทุกทีมรู้ว่าแต่ละ '
                           'requirement มาจากอะไร'},
               {'chapter_id': 'interview-qa',
                'chapter_title': 'Interview Q&A / Key Answers',
                'order': 10,
                'content': 'INTERVIEW Q&A / KEY ANSWERS\n'
                           '\n'
                           '6 KEY SENTENCES TO KEEP IN MIND\n'
                           '\n'
                           '1. I start from customer need, but I always balance it with risk and '
                           'business economics.\n'
                           '\n'
                           '2. I look at the product end to end, not only acquisition.\n'
                           '\n'
                           '3. Growth alone is not success if portfolio quality is not '
                           'sustainable.\n'
                           '\n'
                           '4. I use data to identify the problem, then validate the solution '
                           'before scaling.\n'
                           '\n'
                           '5. For lending, PMF must include both customer fit and risk-adjusted '
                           'profitability.\n'
                           '\n'
                           '6. My role is to connect Product, Risk, Compliance, Operations and '
                           'Technology toward the same business outcome.\n'
                           '\n'
                           '\n'
                           '36) ถ้าคุณได้ตำแหน่งนี้ 90 วันแรกจะทำอะไร?\n'
                           '\n'
                           'Answer:\n'
                           '\n'
                           '30 วันแรก:\n'
                           'Understand portfolio, customers, performance, risk, team และ '
                           'stakeholders\n'
                           '\n'
                           '60 วัน:\n'
                           'Identify key gaps/opportunities และ prioritize\n'
                           '\n'
                           '90 วัน:\n'
                           'Align roadmap และเริ่ม initiative ที่มี clear impact พร้อม KPI\n'
                           '\n'
                           '\n'
                           'SPEAKING PRACTICE\n'
                           '\n'
                           'Answer first, then compare with the suggested structure.'}]},
 {'book_id': 'technical-product-manager-ai-llm',
  'title': 'Technical Product Manager + AI/LLM',
  'subtitle': '30-Day Intensive Bootcamp · Full Book',
  'author': 'Personal Study Library',
  'content_type': 'Book',
  'category': 'Technical Product Management',
  'description': '30-day TPM bootcamp covering software/system foundations, technical delivery, '
                 'AI/ML/LLM, platform product, production readiness, Chinese technical '
                 'communication, and interview readiness.',
  'cover_emoji': '📗',
  'chapters': [{'chapter_id': 'tpm-day-1',
                'chapter_title': 'Day 1 — บทบาท Technical Product Manager',
                'order': 1,
                'content': 'Learning Objective\n'
                           ' เข้าใจว่าตำแหน่งนี้สร้าง value ตรงไหน และแตกต่างจาก PM, PO, BA, '
                           'Project Manager และ Engineer อย่างไร\n'
                           '\n'
                           '1.1 แก่นของงาน\n'
                           'Technical Product Manager (TPM) เป็นเจ้าของ “ปัญหา + ผลลัพธ์ + '
                           'การตัดสินใจเชิงเทคนิคในระดับ Product” โดย\n'
                           'ไม่จำเป็นต้องเป็นคนลงมือ implement เอง สิ่งสำคัญคือสามารถเชื่อม '
                           'Business Goal กับ Architecture, Data, API,\n'
                           'Security, Reliability และ Delivery ได้\n'
                           '\n'
                           '1.2 เส้นแบ่งของบทบาท\n'
                           'PM เน้น Why/What, Engineer เน้น How, TPM ต้องเข้าใจ Why/What และเข้าใจ '
                           'How มากพอที่จะประเมิน trade-off\n'
                           'และถามคำถามถูกจุด ส่วน Project Manager เน้น '
                           'timeline/resource/dependency และ BA เน้น\n'
                           'process/requirement detail แม้ในองค์กรจริงบทบาทจะ overlap กันได้\n'
                           '\n'
                           '1.3 สิ่งที่ TPM ต้องถือในหัวพร้อมกัน\n'
                           'User value, business outcome, technical feasibility, delivery risk, '
                           'security/compliance, performance,\n'
                           'scalability, cost และ operability หลัง go-live '
                           'ทั้งหมดต้องถูกพิจารณาเป็นระบบเดียว ไม่ใช่แยกเป็นฝ่าย ๆ\n'
                           '\n'
                           '1.4 ตัวอย่างจากคำขอ Business\n'
                           'Business บอก “อยากให้ลูกค้าเห็นสถานะสินเชื่อ real-time” TPM '
                           'ต้องถามต่อ: source of truth อยู่ไหน? real-time\n'
                           'หมายถึงกี่วินาที? update event มาจากระบบใด? ถ้า downstream ล่มจะ '
                           'fallback อย่างไร? มี SLA หรือไม่? มี PII อะไร\n'
                           'ถูกส่งผ่าน? metric หลัง launch คืออะไร?\n'
                           '\n'
                           '1.5 Mental model\n'
                           'ทุก feature ให้คิดเป็น 7 ชั้น: User -> Interface -> Service/API -> '
                           'Logic -> Data -> Dependencies -> Operations. ถ้า\n'
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
                           ' เข้าใจการไหลของ request ตั้งแต่หน้าจอผู้ใช้ไปถึง backend และ data '
                           'source\n'
                           '\n'
                           '2.1 Frontend คืออะไร\n'
                           'Frontend คือส่วนที่ user สัมผัส เช่น mobile app, web app, admin portal '
                           'หน้าที่หลักคือรับ input แสดงผล จัดการ\n'
                           'state ฝั่ ง client และเรียก backend ผ่าน network ไม่ควรถือ business '
                           'logic สำคัญทั้งหมดไว้ที่ client เพราะ\n'
                           'แก้ไข/ควบคุม/security ยากกว่า\n'
                           '\n'
                           '2.2 Backend คืออะไร\n'
                           'Backend คือ service ฝั่ ง server ที่รับ request ตรวจสอบสิทธิ์ ประมวลผล '
                           'business logic อ่าน/เขียน database เรียก\n'
                           'downstream system และส่ง response กลับ Frontend หนึ่ง product อาจมี '
                           'backend หลาย service\n'
                           '\n'
                           '2.3 Client-Server flow\n'
                           'ตัวอย่าง: User กด “ดูวงเงิน” -> Frontend ตรวจ input -> ส่ง HTTPS '
                           'request -> API Gateway -> Backend ->\n'
                           'Database/Core system -> Backend แปลงผล -> Frontend render ให้ user '
                           'เห็น\n'
                           '\n'
                           '2.4 State และ Session\n'
                           'บางระบบเป็น stateless: ทุก request มีข้อมูลพอให้ server ทำงานได้เอง '
                           'บางระบบใช้ session/token เพื่อจำว่า user\n'
                           'login แล้ว TPM ต้องรู้ว่า session หมดอายุเมื่อไร, user experience '
                           'เมื่อ token expire เป็นอย่างไร และมี security\n'
                           'implication อะไร\n'
                           '\n'
                           '2.5 Debug mindset\n'
                           'เมื่อหน้าจอผิด อย่าพูดว่า “ระบบพัง” ให้แยก: UI render ผิด? request '
                           'ไม่ออก? API error? backend logic ผิด? data\n'
                           'source ผิด? network timeout? การแยก layer ทำให้ incident triage '
                           'เร็วขึ้น\n'
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
                           ' อ่าน API concept และคุยกับ developer เรื่อง request/response, error '
                           'และ authentication ได้\n'
                           '\n'
                           '3.1 API คือสัญญาการสื่อสาร\n'
                           'API คือ interface/contract '
                           'ที่กำหนดว่าระบบหนึ่งจะขอให้อีกระบบทำอะไรได้อย่างไร Contract ประกอบด้วย '
                           'endpoint,\n'
                           'method, parameters, headers, body, response schema, error code และ '
                           'authentication\n'
                           '\n'
                           '3.2 REST mental model\n'
                           'REST API มักใช้ HTTP methods: GET อ่านข้อมูล, POST สร้าง/สั่ง action, '
                           'PUT แทนที่ resource, PATCH แก้บางส่วน,\n'
                           'DELETE ลบ โดย URI ควรสื่อถึง resource เช่น /customers/123/loans\n'
                           '\n'
                           '3.3 JSON\n'
                           'JSON เป็นรูปแบบ key-value ที่ใช้แลกข้อมูล ตัวอย่าง response มี '
                           'applicationId, status, amount. TPM ควรอ่าน\n'
                           'nested JSON, optional field, null และ data type ได้ เพื่อ review '
                           'contract และ acceptance criteria\n'
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
                           'Authentication = คุณคือใคร เช่น OAuth/OIDC/login/token; Authorization '
                           '= คุณมีสิทธิ์ทำอะไร เช่น\n'
                           'role/permission/scope. Token เช่น JWT อาจมี claims ระบุ '
                           'user/role/expiry แต่ TPM ไม่ต้องเขียน token เอง ต้อง\n'
                           'เข้าใจ lifecycle และ risk\n'
                           '\n'
                           '3.6 Idempotency และ Retry\n'
                           'การ retry POST payment ซ้ำอาจเกิด double charge จึงต้องมี idempotency '
                           'key หรือ design ที่ป้องกัน duplicate.\n'
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
                           ' เข้าใจ relational data, key, join, transaction, source of truth และ '
                           'query พื้นฐาน\n'
                           '\n'
                           '4.1 Table และ Key\n'
                           'Relational database เก็บข้อมูลใน table; row คือ record; column คือ '
                           'attribute. Primary Key ระบุ record ไม่ซ้ำ\n'
                           'เช่น customer_id. Foreign Key ใช้เชื่อม relation เช่น loan.customer_id '
                           '-> customer.customer_id\n'
                           '\n'
                           '4.2 SQL ที่ TPM ควรอ่านได้\n'
                           'SELECT เพื่ออ่าน, WHERE กรอง, JOIN รวมข้อมูลจากหลาย table, GROUP BY '
                           'สรุป, ORDER BY เรียง. ไม่จำเป็นต้อง\n'
                           'เป็น DBA แต่ควร query ข้อมูลเพื่อตรวจ incident/metric เบื้องต้นได้\n'
                           '\n'
                           ' SELECT l.loan_id, l.amount, c.name\n'
                           ' FROM loan l\n'
                           ' JOIN customer c ON c.customer_id = l.customer_id\n'
                           " WHERE l.status = 'ACTIVE'\n"
                           ' ORDER BY l.amount DESC;\n'
                           '\n'
                           '4.3 Transaction และ Consistency\n'
                           'ธุรกรรมบางชุดต้องสำเร็จทั้งหมดหรือไม่สำเร็จเลย เช่น debit/credit. '
                           'แนวคิด ACID ช่วยให้เข้าใจ\n'
                           'atomicity/consistency. TPM ต้องถามว่าถ้าขั้นตอนกลาง fail จะ rollback '
                           'หรือ compensate อย่างไร\n'
                           '\n'
                           '4.4 Source of Truth\n'
                           'ข้อมูลเดียวกันอาจ copy อยู่หลายระบบ ต้องกำหนด authoritative source '
                           'เช่น Customer Master เป็น source of\n'
                           'truth ของชื่อ ส่วน Loan Core เป็น source of truth ของ loan status. '
                           'ถ้าไม่ชัดจะเกิด reconciliation issue\n'
                           '\n'
                           '4.5 Eventual consistency\n'
                           'Distributed systems บางครั้งข้อมูลไม่ตรงกันชั่วคราว เช่น update ใน '
                           'core แล้ว cache/search index ตามหลัง 30\n'
                           'วินาที TPM ต้องแปลง technical behavior เป็น UX ที่ยอมรับได้และ SLA '
                           'ที่ชัด\n'
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
                           'API Gateway เป็นประตูหน้าของ backend APIs ช่วย routing, '
                           'authentication, rate limiting, logging หรือ policy\n'
                           'enforcement. ถ้า gateway ล่ม service หลังบ้านที่ปกติดีก็เข้าถึงไม่ได้\n'
                           '\n'
                           '5.3 Synchronous vs Asynchronous\n'
                           'Sync: caller รอ response เหมาะกับงานต้องรู้ผลทันที; Async: ส่ง '
                           'message/event แล้ว process ภายหลังผ่าน\n'
                           'queue/broker เหมาะกับงานยาวหรือ decouple systems เช่น ส่ง notification '
                           'หลัง approve\n'
                           '\n'
                           ' Synchronous:\n'
                           ' Frontend -> API -> Loan Service -> Core -> Response\n'
                           '\n'
                           ' Asynchronous:\n'
                           ' Loan Service -> Event: LoanApproved -> Queue -> Notification Service\n'
                           '\n'
                           '5.4 Queue / Event\n'
                           'Message queue ช่วย buffer load และ retry; event-driven architecture '
                           'ให้ system publish event เช่น\n'
                           'LoanApproved แล้วระบบอื่น subscribe. TPM ต้องถาม delivery guarantee, '
                           'duplicate event, ordering และ\n'
                           'dead-letter handling ใน use case สำคัญ\n'
                           '\n'
                           '5.5 Cache\n'
                           'Cache เก็บข้อมูลที่เรียกบ่อยเพื่อเร็วขึ้น ลด load แต่เสี่ยงข้อมูล '
                           'stale. Product decision คือยอม stale ได้กี่วินาที/นาที\n'
                           '\n'
                           '5.6 Bottleneck & SPOF\n'
                           'Single Point of Failure คือ component เดียวล่มแล้วระบบหลักหยุด. '
                           'Bottleneck คือจุดที่จำกัด throughput.\n'
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
                           ' เข้าใจเส้นทาง code จาก developer ไป production และศัพท์ DevOps ที่ '
                           'TPM ต้องคุยได้\n'
                           '\n'
                           '6.1 Environment\n'
                           'DEV ใช้พัฒนา, SIT ทดสอบ integration, UAT ให้ business ยืนยัน, '
                           'Staging/Pre-prod ใกล้ production, Production\n'
                           'ใช้งานจริง. ชื่อ environment ต่างกันตามองค์กร แต่หลักคือแยกความเสี่ยง\n'
                           '\n'
                           '6.2 Cloud primitives\n'
                           'Compute รัน application, storage เก็บไฟล์/object, database เก็บ '
                           'structured data, network เชื่อมระบบ, load\n'
                           'balancer กระจาย traffic, autoscaling เพิ่ม/ลด capacity. TPM '
                           'ต้องเข้าใจผลต่อ cost/reliability มากกว่ารายละเอียด\n'
                           'command\n'
                           '\n'
                           '6.3 Git\n'
                           'Git เก็บ version ของ code. Developer ทำ branch -> commit -> pull '
                           'request/merge request -> review ->\n'
                           'merge. TPM ควรรู้ว่า feature ไหนอยู่ branch/version ไหนเพื่อเชื่อม '
                           'roadmap กับ release\n'
                           '\n'
                           '6.4 CI/CD\n'
                           'Continuous Integration รัน build/test อัตโนมัติเมื่อ code เปลี่ยน; '
                           'Continuous Delivery/Deployment ทำ package\n'
                           'และส่งไป environment อย่าง repeatable ลด manual error. Pipeline fail '
                           'ต้องรู้ว่า stage ไหน fail\n'
                           '\n'
                           '6.5 Deployment strategies\n'
                           'Rolling update เปลี่ยนทีละ instance; Blue-Green สลับ environment; '
                           'Canary ปล่อยให้ user กลุ่มเล็กก่อน. TPM\n'
                           'เลือก strategy ตาม risk, traffic, rollback speed\n'
                           '\n'
                           '6.6 Configuration & Secrets\n'
                           'Config ควรแยกตาม environment เช่น endpoint; secrets เช่น password/API '
                           'key ไม่ควร hard-code ใน source\n'
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
                           'Latency, Traffic, Errors, Saturation เป็นสัญญาณหลักที่ใช้ monitor '
                           'service. Product metric และ system metric\n'
                           'ต้องเชื่อมกัน เช่น checkout conversion ลดเพราะ latency เพิ่ม\n'
                           '\n'
                           '7.3 SLA / SLO / SLI\n'
                           'SLI คือสิ่งที่วัด เช่น success rate; SLO คือ target ภายใน เช่น 99.9%; '
                           'SLA คือคำมั่นกับลูกค้า/คู่สัญญาที่อาจมีผลเชิง\n'
                           'พาณิชย์. TPM ต้องระวังอย่าใช้คำสลับกัน\n'
                           '\n'
                           '7.4 Incident process\n'
                           'Detect -> Triage -> Contain -> Communicate -> Diagnose -> Fix -> '
                           'Validate -> Recover -> Postmortem. ระหว่าง\n'
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
                           'Output = feature shipped; Outcome = behavior/business changed. '
                           'ตัวอย่าง: median search time 15 -> 2 นาที,\n'
                           'first-contact resolution 65% -> 85%, incorrect-answer rate <1%\n'
                           '\n'
                           '8.4 Constraints\n'
                           'Regulation, security, legacy dependency, budget, data quality, '
                           'latency, vendor contract ล้วนเป็น constraint\n'
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
                           'PRD ที่ดีควรมี Context, Problem, Objective, Users, Use Cases, Scope, '
                           'Out of Scope, Functional Requirements,\n'
                           'NFR, Data, Dependencies, Risks, Metrics, Rollout, Open Questions\n'
                           '\n'
                           '9.2 Functional requirements\n'
                           'บอกว่าระบบทำอะไร เช่น user ค้น policy, filter ตาม product, เปิด '
                           'citation, ส่ง feedback. แต่ละ requirement ควร\n'
                           'trace ไปที่ user need\n'
                           '\n'
                           '9.3 NFR\n'
                           'Performance (P95 latency), availability, scalability, security, '
                           'privacy, accessibility, auditability,\n'
                           'recoverability, compatibility. คำว่า “เร็ว” “เสถียร” '
                           'ใช้ไม่ได้ถ้าไม่มีตัวเลข/เงื่อนไข\n'
                           '\n'
                           '9.4 Edge cases\n'
                           'Null, duplicate, timeout, partial failure, stale data, permission '
                           'denied, document missing, concurrent\n'
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
                           'รูปแบบ As a [user], I want [capability], so that [outcome] ช่วยย้ำ '
                           'user value แต่ไม่จำเป็นต้องบังคับใช้กับทุก\n'
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
                           'API consumer/provider สามารถใช้ contract tests เพื่อจับ breaking '
                           'change ก่อน integration environment ลด\n'
                           'dependency testing\n'
                           '\n'
                           '11.4 Performance testing\n'
                           'Load test ปริมาณคาดการณ์, stress test เกิน capacity, soak test '
                           'รันนานหา memory leak. TPM ต้องกำหนด\n'
                           'realistic workload และ acceptance\n'
                           '\n'
                           '11.5 Risk-based testing\n'
                           'จัด priority ตาม impact x likelihood: payment duplication, security '
                           'bypass, data corruption ต้อง test ลึกกว่า\n'
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
                           'Estimate ไม่ใช่สัญญาแม่น 100%; งานใหม่/legacy integration มี '
                           'uncertainty สูง. TPM ควรถาม range,\n'
                           'assumptions และ confidence แทนกดให้ทีมตอบเลขเดียว\n'
                           '\n'
                           '12.2 Dependency map\n'
                           'Internal dependency เช่น data team/API platform; external เช่น '
                           'vendor/regulator. ระบุ owner, need-by date,\n'
                           'failure impact และ contingency เพื่อหา critical path\n'
                           '\n'
                           '12.3 MVP\n'
                           'MVP ไม่ใช่ “ทำของแย่” แต่เป็น minimum scope ที่ทดสอบ value/ลด risk '
                           'ได้จริง. ต้องรักษา security/reliability ขั้น\n'
                           'ต่ำ\n'
                           '\n'
                           '12.4 Technical debt\n'
                           'Shortcut ที่ช่วยเร็ววันนี้แต่เพิ่ม cost/risk ในอนาคต เช่น hard-code, '
                           'no tests, duplicated logic. บาง debt ยอมรับได้\n'
                           'ถ้าบันทึกและมี payback plan\n'
                           '\n'
                           '12.5 Trade-off framework\n'
                           'Quality, latency, cost, time, scope, flexibility มักแลกกัน เช่น model '
                           'ใหญ่ quality สูงแต่แพง/ช้า. TPM ต้องทำ\n'
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
                           'Confidentiality: คนไม่มีสิทธิ์ห้ามเห็น, Integrity: '
                           'ข้อมูลห้ามถูกแก้ผิด, Availability: ระบบต้องพร้อมใช้. Product\n'
                           'requirement มักเกี่ยวทั้งสามด้าน\n'
                           '\n'
                           '13.2 Least privilege\n'
                           'ให้ user/service มีสิทธิ์เท่าที่จำเป็นเท่านั้น ใช้ RBAC/ABAC ตามบริบท. '
                           'Admin access ต้อง audit ได้และแยกหน้าที่เมื่อ\n'
                           'จำเป็น\n'
                           '\n'
                           '13.3 Encryption\n'
                           'Data in transit ใช้ TLS/HTTPS; data at rest ควร encrypt ตาม '
                           'sensitivity. TPM ต้องถาม key management และ\n'
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
                           'ต้องคิด mapping, transformation, validation, reconciliation, cutover, '
                           'delta data, rollback/forward-fix,\n'
                           'archive. จำนวน record ตรงไม่พอ ต้องตรวจ business-level integrity\n'
                           '\n'
                           '14.3 Production readiness review\n'
                           'Checklist: capacity, security, monitoring, alerts, dashboards, '
                           'runbook, on-call, backup/restore, DR, vendor\n'
                           'support, SLA, support process, known issues\n'
                           '\n'
                           '14.4 Smoke test\n'
                           'หลัง deployment ทดสอบ critical path สั้น ๆ '
                           'เพื่อยืนยันระบบพื้นฐานทำงานก่อนเปิดเต็ม traffic\n'
                           '\n'
                           '14.5 Rollback vs Forward fix\n'
                           'Rollback เหมาะเมื่อย้อน version ได้ปลอดภัย; schema/data migration '
                           'บางครั้ง rollback ยาก ต้อง forward fix. จึง\n'
                           'ต้องคิดก่อน deploy\n'
                           '\n'
                           '14.6 Hypercare\n'
                           'ช่วงหลัง go-live เพิ่ม monitoring/support cadence และ decision '
                           'threshold เพื่อจับ issue เร็ว\n'
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
                           'AI เป็น umbrella; Machine Learning เรียน pattern จาก data; Deep '
                           'Learning ใช้ neural networks หลายชั้น;\n'
                           'Generative AI สร้าง content; LLM คือ model ภาษา large-scale ที่ทำนาย '
                           'token ถัดไปและเรียน representation\n'
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
                           'Probabilistic output หมายถึงคำตอบเดียวกันอาจแตกต่างและไม่รับประกันถูก '
                           '100%; Product ต้องมี tolerance,\n'
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
                           ' เข้าใจ lifecycle ของ model และ transformer '
                           'แบบไม่ลงคณิตศาสตร์เกินจำเป็น\n'
                           '\n'
                           '16.1 Training\n'
                           'Training ปรับ model parameters จากข้อมูลเพื่อเรียน pattern. Foundation '
                           'model ใช้ compute/data มหาศาล;\n'
                           'enterprise product ส่วนใหญ่ไม่ train foundation model เอง แต่ใช้ '
                           'hosted/open model แล้วต่อยอด\n'
                           '\n'
                           '16.2 Inference\n'
                           'Inference คือการใช้ model ที่ train แล้วตอบ input จริง. ใน production '
                           'TPM สนใจ latency, throughput, tokens,\n'
                           'cost, failure rate, quota และ scaling\n'
                           '\n'
                           '16.3 Transformer mental model\n'
                           'Transformer ใช้ attention เพื่อให้อินพุตแต่ละ token “มอง” token '
                           'อื่นที่เกี่ยวข้อง จึงจับ context ได้ดี. ไม่ต้องคำนวณ\n'
                           'matrix แต่ควรรู้ว่า context length และ tokenization มีผลต่อ '
                           'output/cost\n'
                           '\n'
                           '16.4 Pretraining / Instruction tuning / Alignment\n'
                           'Pretraining เรียนภาษากว้าง ๆ; instruction tuning ช่วยทำตามคำสั่ง; '
                           'alignment techniques ทำให้ตอบสอดคล้อง\n'
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
                           'รวม system prompt + user prompt + retrieved docs + conversation + '
                           'output ต้องอยู่ใน context budget. ใส่\n'
                           'ข้อมูลมากเกินไม่ใช่ดีเสมอ เพราะ noise เพิ่มและแพง\n'
                           '\n'
                           '17.3 Prompt structure\n'
                           'Prompt ที่ดีมี Role/Goal, Context, Rules, Input, Output format, '
                           'Examples และ refusal/fallback behavior ตาม\n'
                           'use case. Structured output เช่น JSON ช่วย downstream parsing\n'
                           '\n'
                           '17.4 Prompt injection\n'
                           'ข้อความในเอกสารหรือ user อาจพยายามเปลี่ยน instruction เช่น “ignore '
                           'previous rules”. Product ต้องแยก\n'
                           'trusted instruction, sanitize/tool permissions และไม่ให้ model '
                           'มีสิทธิ์เกินจำเป็น\n'
                           '\n'
                           '17.5 Prompt versioning\n'
                           'Prompt คือ production artifact ต้อง version, test และ monitor '
                           'เช่นเดียวกับ code เพราะแก้ prompt 1 บรรทัดอาจ\n'
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
                           'Vector search ใช้ similarity metric เช่น cosine similarity เพื่อหา '
                           'chunks ที่ใกล้ query. TPM ไม่ต้องคำนวณเอง แต่\n'
                           'ต้องรู้ว่า top-k และ threshold กระทบ recall/noise\n'
                           '\n'
                           '18.3 Chunking\n'
                           'เอกสารต้องแบ่งเป็น chunks. เล็กเกินสูญ context; ใหญ่เกินดึง noise/ใช้ '
                           'token มาก. ควรใช้ heading/semantic\n'
                           'boundaries และ overlap ตามประเภทเอกสาร\n'
                           '\n'
                           '18.4 Metadata filters\n'
                           'Filter เช่น product, country, effective date, document type ช่วย '
                           'retrieval precision และ policy versioning.\n'
                           'สำคัญใน enterprise knowledge\n'
                           '\n'
                           '18.5 Hybrid search\n'
                           'รวม keyword/BM25 กับ vector search เพื่อจับทั้ง exact terms เช่น '
                           'product code และ semantic meaning\n'
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
                           ' ออกแบบ Retrieval-Augmented Generation และรู้จุด failure ของแต่ละ '
                           'stage\n'
                           '\n'
                           '19.1 RAG flow\n'
                           'Ingestion: อ่านเอกสาร -> clean -> chunk -> embedding -> vector index. '
                           'Query: user question -> optional\n'
                           'rewrite -> retrieve -> rerank/filter -> build prompt -> LLM -> answer '
                           '+ citation -> feedback/logging\n'
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
                           'ถ้าดึงเอกสารผิด ต่อให้ LLM เก่งก็อาจตอบผิด. Evaluation ต้องแยก '
                           'retrieval quality กับ answer generation เพื่อ\n'
                           'debug ได้\n'
                           '\n'
                           '19.4 Freshness & versioning\n'
                           'Policy มี effective date ต้อง ensure index อัปเดตและเลิกใช้ version '
                           'เก่า. Ingestion pipeline ต้อง monitor failed\n'
                           'documents และ stale index\n'
                           '\n'
                           '19.5 Access-aware RAG\n'
                           'User ไม่ควร retrieve เอกสารที่ไม่มีสิทธิ์ แม้ final answer ถูก mask '
                           'ภายหลัง ควร enforce permissions ตั้งแต่\n'
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
                           'ถ้าปัญหาคือ model ไม่รู้ข้อมูลบริษัท/ข้อมูลเปลี่ยนบ่อย ใช้ RAG เพราะ '
                           'update knowledge source ได้โดยไม่ retrain\n'
                           '\n'
                           '20.3 Fine-tuning for behavior\n'
                           'เหมาะเมื่ออยากปรับ style/task pattern/structured behavior จาก examples '
                           'จำนวนมาก แต่ไม่ใช่วิธีดีที่สุดสำหรับ\n'
                           'factual knowledge ที่เปลี่ยนรายวัน\n'
                           '\n'
                           '20.4 Model selection\n'
                           'ดู quality บน use case จริง, latency, cost, context, language, tool '
                           'use, privacy, deployment region, SLA,\n'
                           'vendor lock-in, stability. Benchmark ด้วย evaluation set ของตัวเอง\n'
                           '\n'
                           '20.5 Small vs large models\n'
                           'Model ใหญ่ไม่ได้ชนะทุก use case. Routing อาจส่ง simple task ไป model '
                           'เล็กและ difficult task ไป model ใหญ่ ลด\n'
                           'cost/latency\n'
                           '\n'
                           '20.6 Build vs buy\n'
                           'ซื้อ API เร็วแต่ vendor dependency; self-host ควบคุมสูงแต่ต้องมี '
                           'infra/ops. Decision ต้องรวม total cost of\n'
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
                           'Human review แม่นด้าน nuance แต่แพง/ช้า; automated/LLM-as-judge '
                           'scalable แต่มี bias. ใช้ผสมและ calibrate\n'
                           'กับ human labels\n'
                           '\n'
                           '21.4 Hallucination\n'
                           'คำตอบดูสมเหตุผลแต่ unsupported/ผิด. ลดด้วย RAG, constrained output, '
                           'tool validation, temperature,\n'
                           'refusal when evidence insufficient และ human review แต่ต้องวัดจริง\n'
                           '\n'
                           '21.5 Guardrails\n'
                           'Input/output filters, policy checks, PII masking, tool permission, '
                           'rate limit, safe fallback. Guardrail เป็นชั้น\n'
                           'เพิ่ม ไม่แทน secure architecture\n'
                           '\n'
                           '21.6 Regression\n'
                           'ทุก model/prompt/retrieval change ต้องรัน benchmark เดิมเพื่อดูว่า '
                           'quality ดีขึ้นส่วนหนึ่งแต่พังอีกส่วนหรือไม่\n'
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
                           'ให้คะแนน Value, Feasibility, Data readiness, Risk, Adoption friction. '
                           'Use case ที่ value สูงแต่ risk/ข้อมูลไม่\n'
                           'พร้อมอาจเริ่มด้วย assistive mode ก่อน autonomous\n'
                           '\n'
                           '22.2 Assist vs Automate\n'
                           'Assist: AI เสนอคำตอบให้คนตัดสินใจ; Automate: AI execute action. ยิ่ง '
                           'automation สูงต้องเพิ่ม confidence\n'
                           'threshold, approval, audit และ rollback\n'
                           '\n'
                           '22.3 Human-in-the-loop\n'
                           'กำหนดว่าเคสไหนต้อง human review เช่น low confidence, high amount, '
                           'policy exception, sensitive content.\n'
                           'HITL ต้องออกแบบ queue/SLA ไม่ใช่ใส่ชื่อไว้เฉย ๆ\n'
                           '\n'
                           '22.4 Adoption\n'
                           'Internal AI tool สำเร็จไม่ใช่แค่ model score ต้อง integrate workflow, '
                           'ลด clicks, training, trust/citation,\n'
                           'feedback loop และ champion network\n'
                           '\n'
                           '22.5 North Star + guardrails\n'
                           'North star เช่น time-to-answer ลดลง; guardrail เช่น incorrect-answer '
                           'rate ไม่เกิน 1%, privacy incident = 0\n'
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
                           'User -> UI -> API/Auth -> Orchestrator -> Retrieval/Tools -> Model '
                           'Gateway -> LLM ->\n'
                           'Post-processing/Guardrails -> Response. ด้านข้างมี Logging, '
                           'Evaluation, Cost tracking, Feedback, Secrets,\n'
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
                           'Model timeout, rate limit, retrieval empty, malformed JSON, tool '
                           'error, provider outage. ทุก failure ต้องมี\n'
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
                           'AI cost อาจมาจาก input/output tokens, embedding, vector '
                           'storage/search, reranking, tool APIs, compute,\n'
                           'observability และ human review. อย่าดู token price อย่างเดียว\n'
                           '\n'
                           '24.2 Unit economics\n'
                           'Cost per request x requests/day x days/month = base variable cost. '
                           'เพิ่ม peak capacity, cache hit, retry rate,\n'
                           'long-context users และ support overhead เพื่อ scenario planning\n'
                           '\n'
                           '24.3 Latency budget\n'
                           'แบ่ง latency end-to-end: auth 100ms + retrieval 400ms + rerank 300ms + '
                           'LLM 2.5s + postprocess 200ms =\n'
                           '~3.5s. ถ้าช้าให้ optimize component ที่กิน budget มากที่สุด\n'
                           '\n'
                           '24.4 P50/P95/P99\n'
                           'Average ซ่อน tail latency. P95 = 95% requests เร็วกว่าค่านี้. User '
                           'experience มักเสียจาก tail จึงควรตั้ง SLO เป็น\n'
                           'percentile\n'
                           '\n'
                           '24.5 Caching\n'
                           'Cache embedding/retrieval/response สำหรับ query ซ้ำบางกรณี ลด '
                           'cost/latency แต่ต้องคิด\n'
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
                           ' เข้าใจ lifecycle, versioning, deployment, monitoring, evaluation ตาม '
                           'JD สาย AI platform\n'
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
                           'hosted models อาจ registry configuration/provider version แทน binary '
                           'model\n'
                           '\n'
                           '25.3 Experiment tracking\n'
                           'บันทึกว่า run ไหนใช้ dataset/prompt/model/config อะไรและได้ metric '
                           'เท่าไร เพื่อเปรียบเทียบแบบ reproducible\n'
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
                           'Model/prompt/retriever versions ต้อง trace response ได้. '
                           'ถ้าคุณไม่ตอบได้ว่า “คำตอบนี้มาจาก version ไหน”\n'
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
                           'Platform ให้ capabilities reusable แก่ internal developers/teams เช่น '
                           'API platform, model gateway, feature\n'
                           'store, deployment platform. “User” คือ developer/data '
                           'scientist/product team\n'
                           '\n'
                           '26.2 Self-service\n'
                           'เป้าหมายคือให้ทีมทำงานเองผ่าน portal/API/SDK โดยมี guardrails '
                           'ไม่ต้องเปิด ticket ทุกเรื่อง. Metric เช่น time-to-\n'
                           'first-success, setup lead time, deployment frequency\n'
                           '\n'
                           '26.3 Golden path\n'
                           'สร้างวิธีมาตรฐานที่ง่ายและปลอดภัย เช่น template service + '
                           'logging/security CI pipeline พร้อมใช้ แต่อนุญาต\n'
                           'escape hatch สำหรับ case พิเศษ\n'
                           '\n'
                           '26.4 Developer Experience\n'
                           'Docs, SDK, examples, error messages, sandbox, observability และ '
                           'support เป็น product experience.\n'
                           'Platform ที่ technically powerful แต่ใช้ยาก adoption ต่ำ\n'
                           '\n'
                           '26.5 Multi-tenancy\n'
                           'หลายทีมใช้ platform เดียว ต้องแยก quota, cost, permissions, data, '
                           'environments และ noisy-neighbor risk\n'
                           '\n'
                           '26.6 Platform roadmap\n'
                           'Prioritize common pain across teams, not one-off request '
                           'จากทีมเสียงดัง. ใช้ adoption, reuse, time saved,\n'
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
                           'Business: revenue/cost/risk; Product: adoption/task success/time '
                           'saved; AI: quality/groundedness; System:\n'
                           'latency/availability/error; Cost: unit economics. ต้องเห็น causal '
                           'chain\n'
                           '\n'
                           '27.2 Leading vs Lagging\n'
                           'Leading เช่น weekly active users/feedback; lagging เช่น '
                           'productivity/cost saving. อย่ารอ lagging metric อย่าง\n'
                           'เดียว\n'
                           '\n'
                           '27.3 Experiment design\n'
                           'Define hypothesis, population, control/treatment, primary metric, '
                           'guardrails, duration, sample. AI\n'
                           'experiments ต้อง freeze/version configs เพื่อ interpret results\n'
                           '\n'
                           '27.4 Rollout stages\n'
                           'Internal dogfood -> pilot -> limited GA -> wider rollout -> full. '
                           'แต่ละ stage มี exit criteria เช่น quality,\n'
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
                           'ตัวอย่าง: hallucination เพิ่มหลังเปลี่ยน model, prompt injection ทำให้ '
                           'leak data, retrieval index stale, provider\n'
                           'outage, cost spike, toxic output. Severity ต้องรวม customer/regulatory '
                           'impact\n'
                           '\n'
                           '28.2 Containment\n'
                           'Feature flag off, route fallback model, disable tool action, restrict '
                           'high-risk users, freeze ingestion, revert\n'
                           'prompt/model. ต้องมี runbook ล่วงหน้า\n'
                           '\n'
                           '28.3 Investigation\n'
                           'Trace request -> model/prompt/retrieval/tool version -> logs -> source '
                           'docs -> provider status. ถ้า lineage ไม่\n'
                           'ครบ debug ช้า\n'
                           '\n'
                           '28.4 Governance\n'
                           'Model inventory, use-case risk tier, approval, evaluation evidence, '
                           'data classification, access, change\n'
                           'management, audit, monitoring และ retirement เป็น lifecycle '
                           'governance\n'
                           '\n'
                           '28.5 Responsible AI\n'
                           'Fairness, transparency, privacy, explainability ตาม use case. '
                           'High-stakes decision ควรมี clear\n'
                           'accountability และ human oversight\n'
                           '\n'
                           '28.6 Postmortem\n'
                           'Blameless: timeline, impact, root cause, contributing factors, '
                           'detection gap, corrective actions,\n'
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
                           'เปิด meeting ด้วย objective + decisions needed. ระหว่างคุยถาม fact '
                           'ก่อน opinion: current behavior, expected\n'
                           'behavior, logs, version, environment, reproduction steps, impact. '
                           'ปิดด้วย action/owner/date\n'
                           '\n'
                           '29.2 Clarification language\n'
                           'ใช้ประโยคสั้น ชัด ไม่แปลศัพท์ technical ทุกคำเป็นจีนก็ได้ เพราะ vendor '
                           'จีนใช้ API, JSON, log, bug, release, server\n'
                           'ปนอังกฤษบ่อย เป้าหมายคือสื่อสารงาน ไม่ใช่สอบภาษา\n'
                           '\n'
                           '29.3 Defect discussion\n'
                           'ต้องแยก expected vs actual, reproducible steps, environment/version, '
                           'data, severity, workaround, root\n'
                           'cause, ETA, regression scope\n'
                           '\n'
                           '29.4 Architecture discussion\n'
                           'ถาม data flow, interface, auth, dependency, sync/async, timeout/retry, '
                           'fallback, performance, security,\n'
                           'monitoring\n'
                           '\n'
                           '29.5 Negotiation\n'
                           'เมื่อ ETA ไม่ทัน ให้ถาม option: ลด scope, parallelize, workaround, '
                           'phase delivery, additional resource และ\n'
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
                           ' รวมทุกอย่างเป็น project story และเตรียมตอบ interview Technical '
                           'Product Manager\n'
                           '\n'
                           '30.1 Capstone\n'
                           'ออกแบบ AI Banking Knowledge Assistant: Problem = staff ค้น policy ช้า. '
                           'Users = loan officers. Goal = time-\n'
                           'to-answer <2 นาที, citation accuracy >=98%, incorrect answer <1%, P95 '
                           'latency <5s\n'
                           '\n'
                           '30.2 Architecture\n'
                           'Web UI -> SSO/Auth -> API -> AI Orchestrator -> Access-aware Retrieval '
                           '-> Vector DB -> Reranker -> Model\n'
                           'Gateway -> LLM -> Citation/Guardrails -> Response. '
                           'Logging/evaluation/cost/feedback ครอบทุก request\n'
                           '\n'
                           '30.3 Product decisions\n'
                           'RAG แทน fine-tuning สำหรับ policy freshness, metadata filter ตาม '
                           'product/effective date, human review\n'
                           'สำหรับ high-risk question, canary rollout 100 users, fallback to '
                           'keyword search/manual source\n'
                           '\n'
                           '30.4 Production plan\n'
                           'Evaluation set 500+ cases, security review, load test, runbook, '
                           'model/prompt versioning, monitoring, cost\n'
                           'alert, incident owner, hypercare. เปลี่ยน model ต้อง regression ก่อน '
                           'promote\n'
                           '\n'
                           '30.5 Interview framing\n'
                           'ตอบแบบ Situation -> Problem -> Decision -> Trade-off -> Technical '
                           'depth -> Result -> Learning. อย่าเพียง list\n'
                           'tools; อธิบายว่าทำไมตัดสินใจและ metric เปลี่ยนอย่างไร\n'
                           '\n'
                           '30.6 Gap honesty\n'
                           'ถ้ายังไม่เคย build ML platform จริง ให้พูดตรง ๆ ว่ามี production '
                           'product experience อะไร แล้วแสดง\n'
                           'portfolio/technical understanding ที่เติมมา หลีกเลี่ยงการ claim '
                           'hands-on engineering ที่ไม่ได้ทำ\n'
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
                           '3. ตอบ “How do you manage AI model changes in production?” ภายใน 90 '
                           'วินาที\n'
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
                           '\uf0b7     วาด user -> frontend -> API -> services -> data -> '
                           'downstream\n'
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
                           'API                                                       What makes '
                           'an API change backward incompatible?\n'
                           'Reliability                                               Your P95 '
                           'latency jumped from 2s to 8s after release.\n'
                           'Incident                                                  A downstream '
                           'service is unavailable. How do you\n'
                           'Data                                                      Two systems '
                           'show different customer status. How do\n'
                           'AI                                                        Why would '
                           'you choose RAG instead of fine-tuning?\n'
                           'AI Evaluation                                             How would '
                           'you evaluate hallucination and\n'
                           'LLMOps                                                    How do you '
                           'safely roll out a new model version?\n'
                           'Cost                                                      Quality is '
                           '2% better on a model that costs 4x. How\n'
                           'Platform                                                  How would '
                           'you build an internal model gateway used\n'
                           'Security                                                  How do you '
                           'protect PII in an LLM application?\n'
                           'Product                                                   A '
                           'stakeholder insists on a feature but data shows low\n'
                           'Trade-off                                                 Deadline is '
                           'fixed. Which scope do you cut and how do\n'
                           'Vendor                                                    A Chinese '
                           'vendor says the fix takes 3 weeks. How do you\n'
                           'Leadership                                                Tell me '
                           'about a technical decision where you aligned\n'
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
                           ' Version model/prompt/retrieval config -> run offline regression '
                           'evaluation -> security/quality gate ->\n'
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
                           '         เป้าหมายต่อไป: ลงมือทำ Portfolio 1-2 projects และสะสม '
                           'technical delivery experience จริง'}]}]


# =========================================================
# STATE
# =========================================================
def initialize_state():
    defaults = {
        "page": "Home",
        "books": DEMO_BOOKS,
        "current_book_id": DEMO_BOOKS[0]["book_id"],
        "current_chapter_index": 0,
        "last_read": {
            DEMO_BOOKS[0]["book_id"]: 0,
        },
        "bookmarks": [],
        "font_size": 19,
        "pending_import": None,
        "import_mode": "Paste / Type Text",
        "tts_nonce": 0,
        "reader_notice": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    # Sync built-in books after code updates while preserving true user-imported books.
    # Also remove obsolete built-in versions kept in Streamlit session state.
    builtin_ids = {book["book_id"] for book in DEMO_BOOKS}
    legacy_builtin_ids = {
        "technical-product-manager-ai-llm-chinese",
    }

    current_books = st.session_state.get("books", [])
    imported_books = [
        book for book in current_books
        if book.get("book_id") not in builtin_ids
        and book.get("book_id") not in legacy_builtin_ids
    ]
    st.session_state.books = DEMO_BOOKS + imported_books

    # Migrate an open legacy TPM book to the new no-Chinese version.
    if st.session_state.get("current_book_id") in legacy_builtin_ids:
        st.session_state.current_book_id = "technical-product-manager-ai-llm"
        st.session_state.current_chapter_index = 0

    # Migrate reading progress from the old TPM id if it exists.
    old_tpm_progress = st.session_state.last_read.pop(
        "technical-product-manager-ai-llm-chinese", None
    )
    if old_tpm_progress is not None:
        st.session_state.last_read["technical-product-manager-ai-llm"] = old_tpm_progress

    # Remove bookmarks that point to the obsolete TPM book version.
    st.session_state.bookmarks = [
        bookmark
        for bookmark in st.session_state.get("bookmarks", [])
        if bookmark.get("book_id") not in legacy_builtin_ids
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


def calculate_progress(book, chapter_index: int) -> int:
    total = max(len(book.get("chapters", [])), 1)
    current = min(max(chapter_index + 1, 1), total)
    return round((current / total) * 100)


def open_book(book_id: str, chapter_index=None):
    book = get_book(book_id)
    if not book:
        return
    st.session_state.current_book_id = book_id
    if chapter_index is None:
        chapter_index = st.session_state.last_read.get(book_id, 0)
    chapter_index = max(0, min(chapter_index, max(len(book["chapters"]) - 1, 0)))
    st.session_state.current_chapter_index = chapter_index
    st.session_state.last_read[book_id] = chapter_index
    st.session_state.page = "Reader"


def open_chapter(index: int):
    book = current_book()
    if not book:
        return
    index = max(0, min(index, len(book["chapters"]) - 1))
    st.session_state.current_chapter_index = index
    st.session_state.last_read[book["book_id"]] = index


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
        st.session_state.bookmarks = [
            x for x in st.session_state.bookmarks
            if not (
                x["book_id"] == book["book_id"]
                and x["chapter_id"] == chapter["chapter_id"]
            )
        ]
        st.session_state.reader_notice = "Bookmark removed."
    else:
        st.session_state.bookmarks.append(
            {
                "book_id": book["book_id"],
                "book_title": book["title"],
                "chapter_id": chapter["chapter_id"],
                "chapter_title": chapter["chapter_title"],
                "chapter_index": st.session_state.current_chapter_index,
                "saved_item": chapter["chapter_title"],
            }
        )
        st.session_state.reader_notice = "Saved to Bookmarks."


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


def save_book(title, author, content_type, category, description, raw_text):
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
        "chapters": chapters,
    }

    st.session_state.books.append(book)
    st.session_state.last_read[book_id] = 0
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
            background: rgba(247,241,234,.90);
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


        /* ---------- TOP READER PASTEL-RAINBOW CONTROLS ---------- */
        .st-key-reader_top_actions [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(1) .stButton > button {{
            background: linear-gradient(
                120deg,
                #E2F3FA 0%,
                #C7E3F2 34%,
                #BFD7EA 67%,
                #DADCF3 100%
            ) !important;
            border: 1px solid #AECEE2 !important;
            color: var(--taupe) !important;
            box-shadow: 0 5px 14px rgba(191,215,234,.25) !important;
        }}

        .st-key-reader_top_actions [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(1) .stButton > button:hover {{
            background: linear-gradient(
                120deg,
                #D8EEF8 0%,
                #BFDCEC 34%,
                #B5D1E6 67%,
                #CFD4EF 100%
            ) !important;
            border-color: #9FC3DA !important;
        }}

        .st-key-reader_top_actions [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(2) .stButton > button {{
            background: linear-gradient(
                120deg,
                #F0E8F8 0%,
                #DED4F1 34%,
                #CFC7E8 67%,
                #E7D9EF 100%
            ) !important;
            border: 1px solid #C5B9DF !important;
            color: var(--taupe) !important;
            box-shadow: 0 5px 14px rgba(207,199,232,.28) !important;
        }}

        .st-key-reader_top_actions [data-testid="stHorizontalBlock"] > [data-testid="column"]:nth-child(2) .stButton > button:hover {{
            background: linear-gradient(
                120deg,
                #EADFF5 0%,
                #D6C9EC 34%,
                #C6BCE2 67%,
                #E0CFEB 100%
            ) !important;
            border-color: #B8AAD7 !important;
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
        .nav-wrap {{
            margin: 6px 0 14px;
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
        </style>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# NAVIGATION
# =========================================================
def render_navigation():
    st.markdown('<div class="nav-wrap"></div>', unsafe_allow_html=True)
    cols = st.columns(4, gap="small")
    items = [
        ("Home", "⌂ Home"),
        ("Library", "▤ Library"),
        ("Add", "＋ Add"),
        ("Bookmarks", "♡ Saved"),
    ]
    for col, (page, label) in zip(cols, items):
        with col:
            if st.button(label, key=f"nav_{page}", use_container_width=True):
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

    st.markdown('<div class="section-title">My Library</div>', unsafe_allow_html=True)

    if st.session_state.books:
        for idx, book in enumerate(st.session_state.books[:3]):
            last = st.session_state.last_read.get(book["book_id"], 0)
            render_book_summary(book, last)
            if st.button("Open", key=f"home_open_{idx}_{book['book_id']}"):
                open_book(book["book_id"], last)
                st.rerun()
    else:
        st.caption("No books yet.")

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
def render_library():
    st.markdown("## My Library")

    search = st.text_input(
        "Search",
        placeholder="Search your library...",
        label_visibility="collapsed",
        key="library_search",
    )

    filter_cols = st.columns(4, gap="small")
    filters = ["All", "Books", "Articles", "Notes"]
    if "library_filter" not in st.session_state:
        st.session_state.library_filter = "All"

    for col, label in zip(filter_cols, filters):
        with col:
            if st.button(label, key=f"filter_{label}", use_container_width=True):
                st.session_state.library_filter = label
                st.rerun()

    query = search.strip().lower()
    filtered = []
    for book in st.session_state.books:
        if query and query not in (
            f"{book.get('title','')} {book.get('author','')} "
            f"{book.get('category','')} {book.get('description','')}"
        ).lower():
            continue

        selected = st.session_state.library_filter
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
            """
            <div class="empty-state">
                <div style="font-size:44px;">🐰🌿</div>
                <div class="book-title">No matching books yet.</div>
                <div class="muted">Try another search or add new content.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    for idx, book in enumerate(filtered):
        last = st.session_state.last_read.get(book["book_id"], 0)
        render_book_summary(book, last)
        if st.button("Continue Reading", key=f"lib_open_{idx}_{book['book_id']}"):
            open_book(book["book_id"], last)
            st.rerun()


# =========================================================
# ADD CONTENT / IMPORT
# =========================================================
def render_add_content():
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
        )
        if ok:
            st.session_state.page = "Library"
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
            border-radius: 18px;
            padding: 11px 12px 10px;
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
            width: 36px;
            min-width: 36px;
            height: 36px;
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
            margin-top: 5px;
            padding-left: 2px;
            font-size: 11px;
            color: #8A7468;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        @media (max-width: 430px) {
            .audio-shell {
                padding: 10px 9px 9px;
                border-radius: 16px;
            }

            .top-line {
                gap: 6px;
            }

            .round-btn {
                width: 34px;
                min-width: 34px;
                height: 34px;
                font-size: 15px;
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
    components.html(player_html, height=92, scrolling=False)


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

    top = st.columns([1.15, 1.65, 0.7, 0.7], gap="small")
    with top[0]:
        if st.button("← Back", key="reader_back"):
            st.session_state.page = "Library"
            st.rerun()

    with top[1]:
        st.markdown(
            f'<div style="padding-top:10px;font-size:13px;color:{CI["muted"]};">'
            f'Chapter {idx + 1} / {len(book["chapters"])}</div>',
            unsafe_allow_html=True,
        )

    with top[2]:
        if st.button("Aa", key="font_button"):
            if st.session_state.font_size >= 23:
                st.session_state.font_size = 17
            else:
                st.session_state.font_size += 2
            st.rerun()

    bookmarked = is_bookmarked(book["book_id"], chapter["chapter_id"])
    with top[3]:
        if st.button("♥" if bookmarked else "♡", key="bookmark_current"):
            toggle_bookmark(book, chapter)
            st.rerun()

    # Main reader controls stay at the top so they are reachable without scrolling.
    # Audio becomes a touch-friendly seek bar. Previous / Next remain directly below it.
    with st.container(key="reader_top_actions"):
        render_tts_player(chapter["content"])

        reader_actions = st.columns(2, gap="small")

        with reader_actions[0]:
            prev_disabled = idx <= 0
            if st.button("← Previous", key="prev_chapter", disabled=prev_disabled):
                open_chapter(idx - 1)
                st.rerun()

        with reader_actions[1]:
            next_disabled = idx >= len(book["chapters"]) - 1
            if st.button("Next →", key="next_chapter", disabled=next_disabled):
                open_chapter(idx + 1)
                st.rerun()

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
    st.markdown(
        f"""
        <div class="reader-paper">
            <div class="reader-content" style="font-size:{font_size}px;">
                {escape(chapter["content"])}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    font_cols = st.columns(3, gap="small")
    with font_cols[0]:
        if st.button("A−", key="font_minus"):
            st.session_state.font_size = max(15, st.session_state.font_size - 2)
            st.rerun()
    with font_cols[1]:
        st.markdown(
            f'<div style="text-align:center;padding-top:12px;color:{CI["muted"]};">'
            f'{st.session_state.font_size}px</div>',
            unsafe_allow_html=True,
        )
    with font_cols[2]:
        if st.button("A+", key="font_plus"):
            st.session_state.font_size = min(27, st.session_state.font_size + 2)
            st.rerun()



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
# APP ROUTER
# =========================================================
def main():
    initialize_state()
    inject_css()

    render_navigation()

    page = st.session_state.page
    if page == "Home":
        render_home()
    elif page == "Library":
        render_library()
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

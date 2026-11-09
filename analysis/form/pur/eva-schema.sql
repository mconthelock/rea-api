-- ============================================================================
-- DDL Schema Script for Vendor Evaluation (PUR-EVA / PRO-EVA)
-- Database Target: Oracle / Standard SQL
-- Date: 2026-11-09
-- ============================================================================

-- ============================================================================
-- SCHEMA: WEBFORM (Core Form & Evaluation Tables)
-- ============================================================================

-- 1. Table: WEBFORM.PUREVA_FORM (Main Vendor Evaluation Form)
CREATE TABLE WEBFORM.PUREVA_FORM (
    NFRMNO                NUMBER(4,0)       NOT NULL,
    VORGNO                VARCHAR2(6)       NOT NULL,
    CYEAR                 VARCHAR2(2)       NOT NULL,
    CYEAR2                VARCHAR2(4)       NOT NULL,
    NRUNNO                NUMBER(7,0)       NOT NULL,
    OPERATION             VARCHAR2(1)       NOT NULL,
    VENDCODE              VARCHAR2(10),
    UPSTATUS              VARCHAR2(1)       DEFAULT 'N',
    VENDGROUP             VARCHAR2(50)      NOT NULL,
    VENDPURPOSE           VARCHAR2(50),
    COMNAME               VARCHAR2(200)     NOT NULL,
    VENDTYPE              VARCHAR2(20),
    CONTACT               VARCHAR2(100),
    EMAIL                 VARCHAR2(100),
    WEBSITE               VARCHAR2(200),
    TELNO                 VARCHAR2(50),
    FAX                   VARCHAR2(50),
    BANKNAME              VARCHAR2(100),
    BRANCH                VARCHAR2(100),
    BANKADDR              VARCHAR2(500),
    ACCNUMBER             VARCHAR2(50),
    TERMCODE              VARCHAR2(10),
    CURCODE               VARCHAR2(10),
    COMPLIANCE            VARCHAR2(500),
    COMPLIANCE_OTHER      VARCHAR2(500),
    BUSTYPE_REG           VARCHAR2(200),
    BUSTYPE_SUB           VARCHAR2(200),
    PRODCAT               VARCHAR2(200),
    PRODCAT_OTHER         VARCHAR2(500),
    LEGAL_STATUS          VARCHAR2(50),
    CORPORATE_ID          VARCHAR2(50),
    TAX_ID                VARCHAR2(50),
    CONCERNEDORG          VARCHAR2(50),
    FY_AMOUNT             VARCHAR2(50),
    AMOUNT                NUMBER(12,2)      DEFAULT 0,
    PUR_LEVEL             VARCHAR2(10),
    PUR_STATUS            VARCHAR2(50),
    VENDCAT               VARCHAR2(100),
    CAPITAL               NUMBER(14,2)      DEFAULT 0,
    CAPITAL_CUR           VARCHAR2(10),
    ESTABLISHED           VARCHAR2(50),
    COM_TYPE              VARCHAR2(100),
    COM_OTHER             VARCHAR2(200),
    EMPDIRECT             NUMBER(6,0)       DEFAULT 0,
    EMPINDIRECT           NUMBER(6,0)       DEFAULT 0,
    AVGAGE                VARCHAR2(20),
    QM_STATUS             VARCHAR2(10),
    QM_REASON             VARCHAR2(200),
    CSR_STATUS            VARCHAR2(10),
    CSR_REASON            VARCHAR2(200),
    ENV_STATUS            VARCHAR2(10),
    ENV_REASON            VARCHAR2(200),
    LABOR_STATUS          VARCHAR2(10),
    LABOR_ESTABLISH_DATE  DATE,
    LAND                  NUMBER(10,2)      DEFAULT 0,
    FACTORY               NUMBER(10,2)      DEFAULT 0,
    JUDGEMENT             VARCHAR2(10),
    MJUDGEMENT            VARCHAR2(10),
    ATTACH_TYPE           VARCHAR2(100),
    ATTACH_OTHER          VARCHAR2(500),
    NVFNO                 VARCHAR2(50),
    CONSTRAINT PK_PUREVA_FORM PRIMARY KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO),
    CONSTRAINT FK_PUREVA_FORM_BASE FOREIGN KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO)
        REFERENCES WEBFORM.FORM (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO) ON DELETE CASCADE
);

COMMENT ON TABLE WEBFORM.PUREVA_FORM IS 'ตารางหลักเก็บข้อมูลแบบฟอร์มประเมินผู้ขาย (New Vendor & Annual Evaluation)';
COMMENT ON COLUMN WEBFORM.PUREVA_FORM.OPERATION IS 'ประเภทการประเมิน: N = New Vendor, A = Annual Evaluation';
COMMENT ON COLUMN WEBFORM.PUREVA_FORM.VENDGROUP IS 'กลุ่มผู้ขาย: Direct, 6:Non-Production (6), 8:Sub-Contractor (8)';
COMMENT ON COLUMN WEBFORM.PUREVA_FORM.PUR_LEVEL IS 'ระดับยอดซื้อ: Level A (>=1M), B (>=100K), C (>=10K), D (<10K)';
COMMENT ON COLUMN WEBFORM.PUREVA_FORM.JUDGEMENT IS 'เกรดผลประเมินรวม: A (80+), B (70+), C (60+), D (40+), E (<40)';
COMMENT ON COLUMN WEBFORM.PUREVA_FORM.MJUDGEMENT IS 'เกรดการประเมินตัดสินใจโดย Purchasing Manager (Flow Step 02)';

-- 2. Table: WEBFORM.PUREVA_SCORE (Evaluation Scores by Topic)
CREATE TABLE WEBFORM.PUREVA_SCORE (
    NFRMNO          NUMBER(4,0)       NOT NULL,
    VORGNO          VARCHAR2(6)       NOT NULL,
    CYEAR           VARCHAR2(2)       NOT NULL,
    CYEAR2          VARCHAR2(4)       NOT NULL,
    NRUNNO          NUMBER(7,0)       NOT NULL,
    EVAID           NUMBER(4,0)       NOT NULL,
    TOPIC           VARCHAR2(100)     NOT NULL,
    TOPIC_DESC      VARCHAR2(200),
    SCORE           NUMBER(2,0)       NOT NULL,
    SLEVEL          VARCHAR2(10),
    CONSTRAINT PK_PUREVA_SCORE PRIMARY KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO, EVAID),
    CONSTRAINT FK_PUREVA_SCORE_FORM FOREIGN KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO)
        REFERENCES WEBFORM.PUREVA_FORM (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO) ON DELETE CASCADE
);

COMMENT ON TABLE WEBFORM.PUREVA_SCORE IS 'ตารางคะแนนการประเมินรายหัวข้อ (Price, Order, Service, Quality, Environment etc.)';

-- 3. Table: WEBFORM.PUREVA_PROFIT_TURNOVER (Last 3 Years Turnover & Net Profit)
CREATE TABLE WEBFORM.PUREVA_PROFIT_TURNOVER (
    NFRMNO          NUMBER(4,0)       NOT NULL,
    VORGNO          VARCHAR2(6)       NOT NULL,
    CYEAR           VARCHAR2(2)       NOT NULL,
    CYEAR2          VARCHAR2(4)       NOT NULL,
    NRUNNO          NUMBER(7,0)       NOT NULL,
    ID              NUMBER(4,0)       NOT NULL,
    RECORD_TYPE     VARCHAR2(1)       NOT NULL,
    MYEAR           NUMBER(4,0)       NOT NULL,
    AMOUNT          NUMBER(14,2)      NOT NULL,
    CONSTRAINT PK_PUREVA_TURNOVER PRIMARY KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO, ID),
    CONSTRAINT FK_PUREVA_TURNOVER_FORM FOREIGN KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO)
        REFERENCES WEBFORM.PUREVA_FORM (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO) ON DELETE CASCADE
);

COMMENT ON TABLE WEBFORM.PUREVA_PROFIT_TURNOVER IS 'ตารางผลประกอบการย้อนหลัง 3 ปี (RECORD_TYPE: P = Net Profit/Loss, T = Turnover)';

-- 4. Table: WEBFORM.PUREVA_VENDOR_RELATION (Stakeholders, Customers, Suppliers, Products)
CREATE TABLE WEBFORM.PUREVA_VENDOR_RELATION (
    NFRMNO          NUMBER(4,0)       NOT NULL,
    VORGNO          VARCHAR2(6)       NOT NULL,
    CYEAR           VARCHAR2(2)       NOT NULL,
    CYEAR2          VARCHAR2(4)       NOT NULL,
    NRUNNO          NUMBER(7,0)       NOT NULL,
    ID              NUMBER(4,0)       NOT NULL,
    ENTITY_TYPE     VARCHAR2(1)       NOT NULL,
    ENTITY_NAME     VARCHAR2(200)     NOT NULL,
    PERCENT         NUMBER(3,0)       NOT NULL,
    CONSTRAINT PK_PUREVA_RELATION PRIMARY KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO, ID),
    CONSTRAINT FK_PUREVA_RELATION_FORM FOREIGN KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO)
        REFERENCES WEBFORM.PUREVA_FORM (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO) ON DELETE CASCADE
);

COMMENT ON TABLE WEBFORM.PUREVA_VENDOR_RELATION IS 'ตารางความสัมพันธ์คู่ค้า (ENTITY_TYPE: N = Shareholder, C = Customer, S = Supplier, P = Product)';

-- 5. Table: WEBFORM.PURNVF_ADDRESS (Vendor Address: Shared between NVF and EVA)
CREATE TABLE WEBFORM.PURNVF_ADDRESS (
    NFRMNO          NUMBER(4,0)       NOT NULL,
    VORGNO          VARCHAR2(6)       NOT NULL,
    CYEAR           VARCHAR2(2)       NOT NULL,
    CYEAR2          VARCHAR2(4)       NOT NULL,
    NRUNNO          NUMBER(7,0)       NOT NULL,
    ADDRID          NUMBER(2,0)       NOT NULL,
    ADDRTYPE        VARCHAR2(1)       NOT NULL,
    ADDR1           VARCHAR2(200),
    ADDR2           VARCHAR2(200),
    CITY            VARCHAR2(100),
    STATE           VARCHAR2(100),
    COUNTRY         VARCHAR2(100),
    POSTCODE        VARCHAR2(20),
    CONSTRAINT PK_PURNVF_ADDRESS PRIMARY KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO, ADDRID)
);

COMMENT ON TABLE WEBFORM.PURNVF_ADDRESS IS 'ตารางที่อยู่ผู้ขาย (ADDRTYPE: E = English, T = Thai)';

-- 6. Table: WEBFORM.PUR_FILE (Purchasing System File Attachments)
CREATE TABLE WEBFORM.PUR_FILE (
    NFRMNO          NUMBER(4,0)       NOT NULL,
    VORGNO          VARCHAR2(6)       NOT NULL,
    CYEAR           VARCHAR2(2)       NOT NULL,
    CYEAR2          VARCHAR2(4)       NOT NULL,
    NRUNNO          NUMBER(7,0)       NOT NULL,
    FILE_ID         NUMBER(7,0)       NOT NULL,
    FILE_ONAME      VARCHAR2(255)     NOT NULL,
    FILE_FNAME      VARCHAR2(255)     NOT NULL,
    FILE_PATH       VARCHAR2(500)     NOT NULL,
    FILE_TYPE       NUMBER(2,0)       NOT NULL,
    FILE_USERCREATE VARCHAR2(12)      NOT NULL,
    FILE_DATECREATE DATE              DEFAULT SYSDATE,
    CONSTRAINT PK_PUR_FILE PRIMARY KEY (NFRMNO, VORGNO, CYEAR, CYEAR2, NRUNNO, FILE_ID)
);

COMMENT ON TABLE WEBFORM.PUR_FILE IS 'ตารางไฟล์แนบของฟอร์มจัดซื้อ (FILE_TYPE: 11 = Cert/VAT, 12 = IE, 13 = QA, 2 = Other)';

-- ============================================================================
-- INDEXES FOR PERFORMANCE OPTIMIZATION
-- ============================================================================

CREATE INDEX WEBFORM.IDX_PUREVA_VENDCODE ON WEBFORM.PUREVA_FORM (VENDCODE);
CREATE INDEX WEBFORM.IDX_PUREVA_COMNAME  ON WEBFORM.PUREVA_FORM (COMNAME);
CREATE INDEX WEBFORM.IDX_PUREVA_NVFNO    ON WEBFORM.PUREVA_FORM (NVFNO);
CREATE INDEX WEBFORM.IDX_PUREVA_VENDGRP  ON WEBFORM.PUREVA_FORM (VENDGROUP);

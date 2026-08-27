PRAGMA foreign_keys=ON;
BEGIN;
CREATE TABLE "user" (
  "document_type" TEXT NOT NULL,
  "document_number" TEXT NOT NULL,
  "first_names" TEXT NOT NULL,
  "last_names" TEXT NOT NULL,
  "email" TEXT,
  "phone" TEXT,
  "role" TEXT NOT NULL,
  "status" TEXT NOT NULL DEFAULT 'ACTIVE',
  "password_hash" TEXT NOT NULL,
  "signature_file_name" TEXT,
  "signature_file_path" TEXT,
  "signature_updated_at" TEXT,
  "last_login_at" TEXT,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL,
  "updated_at" TEXT NOT NULL,
  UNIQUE ("document_number"),
  UNIQUE ("email")
);
CREATE TABLE "evidence_category" (
  "code" TEXT NOT NULL,
  "name" TEXT NOT NULL,
  "description" TEXT,
  "icon" TEXT,
  "color" TEXT,
  "sort_order" INTEGER NOT NULL DEFAULT 0,
  "is_active" INTEGER NOT NULL,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL,
  "updated_at" TEXT NOT NULL,
  UNIQUE ("code"),
  UNIQUE ("name")
);
CREATE TABLE "training_group" (
  "created_by" INTEGER NOT NULL,
  "group_number" TEXT NOT NULL,
  "program_name" TEXT NOT NULL,
  "lead_instructor" TEXT,
  "followup_instructor" TEXT,
  "municipality" TEXT,
  "program_level" TEXT,
  "modality" TEXT,
  "sofia_group_status" TEXT,
  "group_validity" TEXT,
  "group_start_date" TEXT,
  "training_end_date" TEXT,
  "ep_start_date" TEXT,
  "apprentices_statistics" TEXT,
  "apprentices_training" TEXT,
  "apprentices_enabled" TEXT,
  "apprentices_rap_pending" TEXT,
  "apprentices_practice" TEXT,
  "apprentices_without_alternative" TEXT,
  "apprentices_certified" TEXT,
  "productive_modalities" TEXT,
  "learning_contract" TEXT,
  "internship" TEXT,
  "productive_project" TEXT,
  "employment_link" TEXT,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL,
  "updated_at" TEXT NOT NULL,
  FOREIGN KEY ("created_by") REFERENCES "user"("id"),
  UNIQUE ("group_number")
);
CREATE TABLE "apprentice" (
  "created_by" INTEGER NOT NULL,
  "student_user_id" INTEGER,
  "group_id" INTEGER,
  "group_number" TEXT NOT NULL,
  "document_type" TEXT NOT NULL,
  "document_number" TEXT NOT NULL,
  "first_names" TEXT NOT NULL,
  "last_names" TEXT NOT NULL,
  "gender" TEXT,
  "phone" TEXT,
  "email" TEXT,
  "municipality_origin" TEXT,
  "program_name" TEXT,
  "program_level" TEXT,
  "group_validity" TEXT,
  "lead_instructor" TEXT,
  "followup_instructor" TEXT,
  "followup_instructor_email" TEXT,
  "ep_modality" TEXT,
  "sofia_status" TEXT,
  "practice_start_date" TEXT,
  "practice_end_date" TEXT,
  "followup_moment1_start" TEXT,
  "followup_moment1_end" TEXT,
  "followup_moment2_start" TEXT,
  "followup_moment2_end" TEXT,
  "followup_moment3_start" TEXT,
  "followup_moment3_end" TEXT,
  "followup_moment4_start" TEXT,
  "followup_moment4_end" TEXT,
  "company_name" TEXT,
  "company_municipality" TEXT,
  "company_address" TEXT,
  "coformador_name" TEXT,
  "coformador_email" TEXT,
  "coformador_phone" TEXT,
  "arl_responsible" TEXT,
  "continues_company" TEXT,
  "individual_management" TEXT,
  "followup_moments" TEXT,
  "evaluation_date" TEXT,
  "english_results" TEXT,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL,
  "updated_at" TEXT NOT NULL,
  FOREIGN KEY ("created_by") REFERENCES "user"("id"),
  FOREIGN KEY ("student_user_id") REFERENCES "user"("id"),
  FOREIGN KEY ("group_id") REFERENCES "training_group"("id"),
  UNIQUE ("document_number")
);
CREATE TABLE "evidence_template" (
  "category_id" INTEGER NOT NULL,
  "code" TEXT NOT NULL,
  "title" TEXT NOT NULL,
  "description" TEXT,
  "allowed_extensions" TEXT,
  "max_file_size_mb" INTEGER,
  "requires_signature" INTEGER NOT NULL,
  "is_required" INTEGER NOT NULL,
  "sort_order" INTEGER NOT NULL DEFAULT 0,
  "is_active" INTEGER NOT NULL,
  "created_by_id" INTEGER,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL,
  "updated_at" TEXT NOT NULL,
  FOREIGN KEY ("category_id") REFERENCES "evidence_category"("id"),
  FOREIGN KEY ("created_by_id") REFERENCES "user"("id"),
  UNIQUE ("code")
);
CREATE TABLE "evidence_activity" (
  "group_id" INTEGER NOT NULL,
  "template_id" INTEGER,
  "category_id" INTEGER,
  "code" TEXT,
  "title" TEXT NOT NULL,
  "description" TEXT,
  "due_start" TEXT,
  "due_end" TEXT,
  "allowed_extensions" TEXT,
  "max_file_size_mb" INTEGER,
  "requires_signature" INTEGER NOT NULL,
  "is_required" INTEGER NOT NULL,
  "is_visible" INTEGER NOT NULL,
  "is_default" INTEGER NOT NULL,
  "origin" TEXT NOT NULL DEFAULT 'template',
  "sort_order" INTEGER NOT NULL DEFAULT 0,
  "created_by_id" INTEGER,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL,
  "updated_at" TEXT NOT NULL,
  FOREIGN KEY ("group_id") REFERENCES "training_group"("id"),
  FOREIGN KEY ("template_id") REFERENCES "evidence_template"("id"),
  FOREIGN KEY ("category_id") REFERENCES "evidence_category"("id"),
  FOREIGN KEY ("created_by_id") REFERENCES "user"("id")
);
CREATE TABLE "evidence_submission" (
  "activity_id" INTEGER NOT NULL,
  "apprentice_id" INTEGER NOT NULL,
  "status" TEXT NOT NULL DEFAULT 'no_entregado',
  "file_name" TEXT,
  "file_path" TEXT,
  "mime_type" TEXT,
  "file_size_bytes" INTEGER,
  "uploaded_at" TEXT,
  "reviewed_at" TEXT,
  "reviewed_by" INTEGER,
  "approved_at" TEXT,
  "approved_by_id" INTEGER,
  "signed_file_name" TEXT,
  "signed_file_path" TEXT,
  "signed_at" TEXT,
  "version_number" INTEGER NOT NULL DEFAULT 1,
  "attempt_number" INTEGER NOT NULL DEFAULT 1,
  "is_latest" INTEGER NOT NULL,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL,
  "updated_at" TEXT NOT NULL,
  FOREIGN KEY ("activity_id") REFERENCES "evidence_activity"("id"),
  FOREIGN KEY ("apprentice_id") REFERENCES "apprentice"("id"),
  FOREIGN KEY ("reviewed_by") REFERENCES "user"("id"),
  FOREIGN KEY ("approved_by_id") REFERENCES "user"("id")
);
CREATE TABLE "evidence_comment" (
  "submission_id" INTEGER NOT NULL,
  "author_id" INTEGER,
  "comment" TEXT NOT NULL,
  "is_internal" INTEGER NOT NULL,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL,
  "updated_at" TEXT NOT NULL,
  FOREIGN KEY ("submission_id") REFERENCES "evidence_submission"("id"),
  FOREIGN KEY ("author_id") REFERENCES "user"("id")
);
CREATE TABLE "evidence_submission_attempt" (
  "submission_id" INTEGER NOT NULL,
  "attempt_number" INTEGER NOT NULL,
  "version_number" INTEGER NOT NULL,
  "status" TEXT NOT NULL DEFAULT 'pendiente_revision',
  "file_name" TEXT,
  "file_path" TEXT,
  "mime_type" TEXT,
  "file_size_bytes" INTEGER,
  "uploaded_at" TEXT,
  "reviewed_at" TEXT,
  "reviewed_by" INTEGER,
  "approved_at" TEXT,
  "approved_by_id" INTEGER,
  "signed_file_name" TEXT,
  "signed_file_path" TEXT,
  "signed_at" TEXT,
  "id" INTEGER PRIMARY KEY,
  "created_at" TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updated_at" TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY ("submission_id") REFERENCES "evidence_submission"("id"),
  FOREIGN KEY ("reviewed_by") REFERENCES "user"("id"),
  FOREIGN KEY ("approved_by_id") REFERENCES "user"("id")
);
CREATE UNIQUE INDEX "uq_evidence_submission_latest_per_activity_apprentice" ON "evidence_submission" ("activity_id","apprentice_id") WHERE is_latest = 1;
CREATE INDEX "ix_evidence_submission_attempt_submission_id" ON "evidence_submission_attempt" ("submission_id");
CREATE INDEX "ix_evidence_submission_attempt_status" ON "evidence_submission_attempt" ("status");

INSERT INTO evidence_category (code,name,color,sort_order,is_active,created_at,updated_at) VALUES
('initial_requirements','Requisitos Iniciales','#39a900',10,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP),
('logs','Bitácoras','#ff6b00',20,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP),
('followup_moments','Momentos de Seguimiento','#00a9c7',30,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP),
('certification_requirements','Requisitos de Certificación','#e3267f',40,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP);

INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'RI-F165','F-165 Formato seleccion modificacion alternativa etapa productiva',0,1,1,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='initial_requirements';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'RI-ARL','Certificado de afiliacion ARL',0,1,2,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='initial_requirements';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'LOG-01','Bitacora 1',0,1,10,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='logs';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'LOG-02','Bitacora 2',0,1,11,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='logs';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'LOG-03','Bitacora 3',0,1,12,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='logs';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'LOG-04','Bitacora 4',0,1,13,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='logs';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'LOG-05','Bitacora 5',0,1,14,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='logs';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'LOG-06','Bitacora 6',0,1,15,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='logs';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'FUP-01','Momento 1: planeacion de la etapa productiva',0,1,20,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='followup_moments';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'FUP-02','Momento 2: seguimiento de la etapa productiva',0,1,21,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='followup_moments';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'FUP-03','Momento 3: evaluacion de la etapa productiva',0,1,22,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='followup_moments';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'FUP-04','Momento 4: adicional (opcional)',0,0,23,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='followup_moments';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'CERT-ID','Copia de documento de Identidad',0,1,30,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='certification_requirements';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'CERT-ICFES','Certificado de presentacion Pruebas ICFES TyT',0,1,31,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='certification_requirements';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'CERT-APE','Certificado de la APE',0,1,32,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='certification_requirements';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'CERT-CARNET','Carnet Destruido',0,1,33,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='certification_requirements';
INSERT INTO evidence_template (category_id,code,title,requires_signature,is_required,sort_order,is_active,created_at,updated_at)
SELECT id,'CERT-COFORMADOR','Certificado de ente coformador aprobando finalizacion de practicas',0,1,34,1,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP FROM evidence_category WHERE code='certification_requirements';

COMMIT;

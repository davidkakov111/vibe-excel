-- CreateTable
CREATE TABLE "SpreadsheetVersion" (
    "id" TEXT NOT NULL,
    "spreadsheetId" TEXT NOT NULL,
    "revision" INTEGER NOT NULL,
    "workbook" JSONB NOT NULL,
    "createdById" TEXT NOT NULL,
    "authorName" TEXT,
    "label" TEXT NOT NULL DEFAULT 'Auto-save',
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "SpreadsheetVersion_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "SpreadsheetVersion_spreadsheetId_createdAt_idx" ON "SpreadsheetVersion"("spreadsheetId", "createdAt" DESC);

-- AddForeignKey
ALTER TABLE "SpreadsheetVersion" ADD CONSTRAINT "SpreadsheetVersion_spreadsheetId_fkey" FOREIGN KEY ("spreadsheetId") REFERENCES "Spreadsheet"("id") ON DELETE CASCADE ON UPDATE CASCADE;
-- CreateEnum
CREATE TYPE "SharePermission" AS ENUM ('VIEW', 'EDIT');

-- AlterTable
ALTER TABLE "Spreadsheet" ADD COLUMN "revision" INTEGER NOT NULL DEFAULT 0;

-- CreateTable
CREATE TABLE "SpreadsheetShare" (
    "id" TEXT NOT NULL,
    "spreadsheetId" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "permission" "SharePermission" NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "SpreadsheetShare_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "SpreadsheetShare_spreadsheetId_userId_key" ON "SpreadsheetShare"("spreadsheetId", "userId");

-- CreateIndex
CREATE INDEX "SpreadsheetShare_userId_idx" ON "SpreadsheetShare"("userId");

-- AddForeignKey
ALTER TABLE "SpreadsheetShare" ADD CONSTRAINT "SpreadsheetShare_spreadsheetId_fkey" FOREIGN KEY ("spreadsheetId") REFERENCES "Spreadsheet"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SpreadsheetShare" ADD CONSTRAINT "SpreadsheetShare_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User"("id") ON DELETE CASCADE ON UPDATE CASCADE;
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
'use server';

/**
 * app/actions/mis-report-actions.ts
 * ----------------------------------
 * All MIS-module Server Actions live here.
 *
 * ## RBAC contract (per directive)
 *   - `generateReport` NEVER throws an unhandled error for access denial.
 *     On `MISAccessDeniedError` it returns a safe object so the calling
 *     Server Component can pass it to `UniversalReportShell` which then
 *     renders the polite <AccessDeniedState> UI.
 *   - `exportReportToExcel` is called from a Client Component button; on
 *     access denial it throws a user-friendly string so ExportExcelButton
 *     can display its `error` state.
 *   - `listCatalogue` already filtered by `requiredPermission`; RBAC is now
 *     enforced end-to-end because `runReport` calls `assertReportAccess`.
 *
 * ## Session note
 *   `getSession()` is a mock that reads the first User from the DB and derives
 *   MIS permissions from `User.role` via `getMISPermissions()`. Once a real
 *   auth provider (NextAuth, JWT, etc.) is wired, replace `getSession` only —
 *   no other function in this file needs to change.
 */

import { prisma } from '@/backend/db';
import { runReport, REGISTRY } from '@/lib/mis/runner';
import {
  dailyRevenueReport,
  billingDetailReport,
  billingItemDetailReport,
  billingSummaryReport,
  billingSummaryDetailReport,
  billingPaymentModeReport,
  billingUhidAdvanceReport,
  billingAdmissionAdvanceReport,
  billingDiscountSummaryReport,
  billingDueSettledReport,
  billingRefundReport,
  billingOpRefundReport,
  billingDateWiseCashReport,
  billingDoctorPayoutReport,
  billingDoctorAccountPayableReport,
  billingIpPackageReport,
  billingHealthCheckupCountReport,
  billingPayerAgreementExpiryReport,
  billingDepositRefundReport,
  billingPendingBillsReport,
  billingPaymentServiceTypeReport,
  billingPaymentSummaryReport,
  billingRevenueSummaryReport,
  billingCancelBillReport,
  billingServiceTypeSummaryReport
} from '@/lib/mis/registry/billing';
import {
  revenueDepartmentWiseReport,
  revenuePayerTypeWiseReport,
  revenuePayerNameWiseReport,
  revenueServiceTypeWiseReport,
  revenueBillingCategoryWiseReport,
  revenueWardWiseReport
} from '@/lib/mis/registry/revenue';
import {
  preRegistrationReport,
  registrationConvertReport,
  registrationReport,
  appointmentReport,
  doctorEventOffReport,
  doctorEventOffSummaryReport,
  appointmentTatReport,
  doctorFootfallReport,
  ipPatientReport,
  ipConversionReport,
  ipCancelReport,
  ipDischargeReport,
  erToIpConversionReport,
  bedStatusReport,
  admissionsListReport,
  emergencyAdmissionsReport,
  bedOccupancyReport,
  emergencyDischargeReport,
  bedTransferReport,
  doctorTransferReport,
  expiredPatientsReport,
  counsellingSummaryReport,
  dischargeTatReport
} from '@/lib/mis/registry/frontdesk';
import {
  diagCardiologyAppointmentReport,
  diagNeurologyAppointmentReport,
  diagNuclearMedicineAppointmentReport,
  diagPathologyAppointmentReport,
  diagPulmonologyAppointmentReport,
  diagRadiologyAppointmentReport,
  diagCardiologyServiceReport,
  diagNeurologyServiceReport,
  diagPulmonologyServiceReport,
  diagPathologyServiceReport,
  diagTatReport,
  diagPathologyTatReport,
  diagRadiologyTatReport,
  diagRadiologyServiceReport
} from '@/lib/mis/registry/diagnostic';
import {
  pharmacyIpIssueReport,
  pharmacyIpItemDetailReport,
  pharmacyOpItemDetailReport,
  pharmacyOpSummaryDetailReport,
  pharmacyIpDailyReport,
  pharmacyDailyIpReturnReport,
  pharmacyDailyIpIssueReport,
  pharmacyOpTaxSummaryReport,
  pharmacyOpTaxDetailsReport,
  pharmacyIpIssueWithTagsReport,
  pharmacyOpSaleWithTagsReport,
  pharmacyItemReorderLevelReport,
  pharmacySupplierListReport,
  pharmacyCurrentStockReport,
  pharmacyExpiryReport,
  pharmacyDoctorWiseSaleReport,
  pharmacyPatientPullOffReport,
  pharmacyDoctorPullOffReport,
  pharmacyIpPullOffReport,
  pharmacyOpSummaryReport,
  pharmacyOpSaleDetailReport,
  pharmacyOpReturnDetailReport,
  pharmacyHsnSummaryReport,
  pharmacySettlementReport,
  pharmacyDoctorWiseDetailReport
} from '@/lib/mis/registry/pharmacy';
import {
  inventoryStockReport,
  inventoryItemWiseStockReport,
  inventoryGrnSummaryReport,
  inventoryGrnReturnSummaryReport,
  inventoryItemMasterReport,
  inventoryGrnDetailReport,
  inventoryGrnReturnDetailReport,
  inventoryStoreToStoreIssueReport,
  inventoryBatchInflowOutflowReport,
  inventoryAsOnDateStockReport,
  inventoryHsnTaxSummaryReport,
  inventoryBinCardBatchReport,
  inventoryBinCardItemReport,
  inventoryMovingItemsReport,
  inventoryGrnPendingCnReport,
  inventoryStoreConsumptionReport
} from '@/lib/mis/registry/inventory';
import {
  otBookingDetailsReport,
  otSurgeryDetailsReport,
  otSurgeryTatReport,
  ambulanceOrdersReport,
  ambulanceRequestReport,
  ambulanceTatReport,
  opticalItemBillingReport,
  opticalProductBillingReport,
  opticalDailySettlementReport,
  opticalDailySettlementSumReport,
  opticalPaymentReport
} from '@/lib/mis/registry/specialized';
import { generateExcelBuffer } from '@/lib/mis/exporter';
import { GenerateReportResponse, JobStatusResponse, ExportExcelResponse } from '@/lib/mis/action-types';
import { getMISPermissions, MISAccessDeniedError } from '@/lib/mis/rbac';
import type { ColumnSpec } from '@/lib/mis/types';

// ─── Session ──────────────────────────────────────────────────────────────────

/**
 * Mock session resolver.
 *
 * Reads the first User from the DB and derives MIS permissions from their
 * `User.role` string via `getMISPermissions()`.
 *
 * IMPORTANT: Replace this function's body with a real auth lookup (NextAuth
 * `getServerSession`, JWT decode, etc.) before production. The public API of
 * `getSession()` must remain unchanged — callers only use `orgId`, `userId`,
 * and `permissions`.
 */
import { getSession as getRealSession } from '@/app/lib/session';

async function getSession() {
  const session = await getRealSession();

  if (!session) {
    throw new Error('UNAUTHORIZED');
  }

  return {
    orgId: session.organization_id,
    userId: session.id,
    permissions: getMISPermissions(session.role),
  };
}

// ─── Catalogue ────────────────────────────────────────────────────────────────

export async function listCatalogue() {
  const session = await getSession();

  const allReports = [
    dailyRevenueReport,
    billingDetailReport,
    billingItemDetailReport,
    billingSummaryReport,
    billingSummaryDetailReport,
    billingPaymentModeReport,
    billingUhidAdvanceReport,
    billingAdmissionAdvanceReport,
    billingDiscountSummaryReport,
    billingDueSettledReport,
    billingRefundReport,
    billingOpRefundReport,
    billingDateWiseCashReport,
    billingDoctorPayoutReport,
    billingDoctorAccountPayableReport,
    billingIpPackageReport,
    billingHealthCheckupCountReport,
    billingPayerAgreementExpiryReport,
    billingDepositRefundReport,
    billingPendingBillsReport,
    billingPaymentServiceTypeReport,
    billingPaymentSummaryReport,
    billingRevenueSummaryReport,
    billingCancelBillReport,
    billingServiceTypeSummaryReport,
    revenueDepartmentWiseReport,
    revenuePayerTypeWiseReport,
    revenuePayerNameWiseReport,
    revenueServiceTypeWiseReport,
    revenueBillingCategoryWiseReport,
    revenueWardWiseReport,
    preRegistrationReport,
    registrationConvertReport,
    registrationReport,
    appointmentReport,
    doctorEventOffReport,
    doctorEventOffSummaryReport,
    appointmentTatReport,
    doctorFootfallReport,
    ipPatientReport,
    ipConversionReport,
    ipCancelReport,
    ipDischargeReport,
    erToIpConversionReport,
    bedStatusReport,
    admissionsListReport,
    emergencyAdmissionsReport,
    bedOccupancyReport,
    emergencyDischargeReport,
    bedTransferReport,
    doctorTransferReport,
    expiredPatientsReport,
    counsellingSummaryReport,
    dischargeTatReport,
    diagCardiologyAppointmentReport,
    diagNeurologyAppointmentReport,
    diagNuclearMedicineAppointmentReport,
    diagPathologyAppointmentReport,
    diagPulmonologyAppointmentReport,
    diagRadiologyAppointmentReport,
    diagCardiologyServiceReport,
    diagNeurologyServiceReport,
    diagPulmonologyServiceReport,
    diagPathologyServiceReport,
    diagTatReport,
    diagPathologyTatReport,
    diagRadiologyTatReport,
    pharmacyIpIssueReport,
    pharmacyIpItemDetailReport,
    pharmacyOpItemDetailReport,
    pharmacyOpSummaryDetailReport,
    pharmacyIpDailyReport,
    pharmacyDailyIpReturnReport,
    pharmacyDailyIpIssueReport,
    pharmacyOpTaxSummaryReport,
    pharmacyOpTaxDetailsReport,
    pharmacyIpIssueWithTagsReport,
    pharmacyOpSaleWithTagsReport,
    pharmacyItemReorderLevelReport,
    pharmacySupplierListReport,
    pharmacyCurrentStockReport,
    pharmacyExpiryReport,
    pharmacyDoctorWiseSaleReport,
    pharmacyPatientPullOffReport,
    pharmacyDoctorPullOffReport,
    pharmacyIpPullOffReport,
    inventoryStockReport,
    inventoryItemWiseStockReport,
    inventoryGrnSummaryReport,
    inventoryGrnReturnSummaryReport,
    inventoryItemMasterReport,
    inventoryGrnDetailReport,
    inventoryGrnReturnDetailReport,
    inventoryStoreToStoreIssueReport,
    inventoryBatchInflowOutflowReport,
    inventoryAsOnDateStockReport,
    inventoryHsnTaxSummaryReport,
    inventoryBinCardBatchReport,
    inventoryBinCardItemReport,
    inventoryMovingItemsReport,
    inventoryGrnPendingCnReport,
    inventoryStoreConsumptionReport,
    otBookingDetailsReport,
    otSurgeryDetailsReport,
    otSurgeryTatReport,
    ambulanceOrdersReport,
    ambulanceRequestReport,
    ambulanceTatReport,
    opticalItemBillingReport,
    opticalProductBillingReport,
    opticalDailySettlementReport,
    opticalDailySettlementSumReport,
    opticalPaymentReport,
    pharmacyOpSummaryReport,
    pharmacyOpSaleDetailReport,
    pharmacyOpReturnDetailReport,
    pharmacyHsnSummaryReport,
    pharmacySettlementReport,
    pharmacyDoctorWiseDetailReport,
    diagRadiologyServiceReport
  ];

  // Only surface reports the session's role can actually run.
  const accessibleReports = allReports.filter((r) =>
    session.permissions.includes(r.requiredPermission)
  );

  const grouped = accessibleReports.reduce((acc, report) => {
    if (!acc[report.category]) acc[report.category] = [];
    // Strip server-only fields before sending to the client.
    // queryFn and countFn are functions — they cannot cross the Server Action
    // boundary and must never be serialised to the client.
    // drillDownTo, drillDownKey, and chartSpec are plain serialisable values;
    // they are intentionally kept so the Hub UI can display drill-down badges
    // and chart previews without needing a separate round-trip.
    // Gap #14 fix: removed drillDownTo from the destructured strip list.
    const { queryFn, countFn, ...clientDef } = report;
    acc[report.category].push(clientDef);
    return acc;
  }, {} as Record<string, any[]>);

  return grouped;
}

// ─── Generate Report ──────────────────────────────────────────────────────────

/**
 * Runs a report and returns a `GenerateReportResponse`.
 *
 * ## Access Denied — safe return instead of throw
 * Per directive: this action MUST NOT throw an unhandled error on access
 * denial. Doing so would cause Next.js to crash to the error boundary from the
 * Server Component page. Instead, on `MISAccessDeniedError` we return:
 *
 *   { async: false, error: 'UNAUTHORIZED' }
 *
 * `UniversalReportShell` checks for `payload.error === 'UNAUTHORIZED'` and
 * renders the polite <AccessDeniedState> UI.
 *
 * All other unexpected errors are still re-thrown so the error boundary
 * (or a wrapping try/catch in page.tsx) can handle them.
 */
export async function generateReport(
  reportId: string,
  filters: unknown
): Promise<GenerateReportResponse> {
  const session = await getSession();

  try {
    const result = await runReport(
      reportId,
      filters,
      session.orgId,
      session.userId,
      session.permissions
    );
    return result;
  } catch (error: any) {
    // ── Safe path: access denial → return, not throw ──────────────────────
    if (error?.code === 'MIS_ACCESS_DENIED') {
      return {
        async: false,
        error: 'UNAUTHORIZED',
      };
    }
    // ── All other errors: re-throw for Next.js error boundary ─────────────
    throw new Error(error.message ?? 'Failed to generate report');
  }
}

// ─── Job Status ───────────────────────────────────────────────────────────────

export async function getJobStatus(jobId: string): Promise<JobStatusResponse> {
  const session = await getSession();

  const job = await prisma.reportJob.findUnique({
    where: { id: jobId },
  });

  // Gap #9 fix: enforce BOTH multi-tenant isolation (organizationId) AND
  // intra-org job ownership (requested_by).
  //
  // Without the requested_by check, any authenticated user within the same
  // organisation could poll another user's job by knowing or guessing its UUID.
  // The two conditions are deliberately kept as a single guard so the error
  // message does not reveal whether the job exists at all (avoids enumeration).
  if (
    !job ||
    job.organizationId !== session.orgId ||
    job.requested_by !== session.userId
  ) {
    throw new Error('Job not found');
  }

  // Gap #17 fix: surface expiry as a first-class status.
  //
  // If a Completed job's expires_at has passed, the file_key URL is no longer
  // valid (the local file has been deleted by the cleanup cron, or the S3
  // pre-signed URL period has lapsed). We return status 'Expired' rather than
  // 'Completed' so ExportExcelButton can render a "Regenerate" prompt instead
  // of a broken download link.
  //
  // Only Completed jobs can expire — Running/Queued/Failed jobs do not have
  // a download file, so expires_at is irrelevant for them.
  const isExpired =
    job.status === 'Completed' &&
    job.expires_at !== null &&
    job.expires_at < new Date();

  return {
    id: job.id,
    status: isExpired ? 'Expired' : job.status,
    progress: job.progress,
    file_key: isExpired ? null : job.file_key,  // do not return a dead URL
    error: isExpired
      ? 'This export has expired. Please regenerate the report.'
      : job.error,
    createdAt: job.createdAt,
    finished_at: job.finished_at,
  };
}

export async function listJobs(): Promise<JobStatusResponse[]> {
  const session = await getSession();

  // Gap #9 fix: scope to the requesting user's OWN jobs only.
  // Previously this fetched all jobs for the organisation, meaning any org
  // member could inspect every other member's export history and file keys.
  //
  // Gap #17 fix: exclude expired Completed jobs from the list.
  //   A job is expired when: status === 'Completed' AND expires_at < now.
  //   We exclude them here with a Prisma OR filter:
  //     Show the job if:
  //       (a) it is NOT Completed (still Queued / Running / Failed — always
  //           visible regardless of expires_at), OR
  //       (b) it IS Completed AND expires_at is either null (never expires)
  //           or still in the future.
  //   This keeps the list clean — no stale download buttons for the user.
  const now = new Date();
  const jobs = await prisma.reportJob.findMany({
    where: {
      organizationId: session.orgId,
      requested_by: session.userId,
      OR: [
        { status: { not: 'Completed' } },                         // (a) non-terminal
        { status: 'Completed', expires_at: null },                // (b-i) completed, no expiry
        { status: 'Completed', expires_at: { gt: now } },         // (b-ii) completed, still fresh
      ],
    },
    orderBy: { createdAt: 'desc' },
    take: 20,
  });

  return jobs.map((job) => ({
    id: job.id,
    status: job.status,
    progress: job.progress,
    file_key: job.file_key,
    error: job.error,
    createdAt: job.createdAt,
    finished_at: job.finished_at,
  }));
}

/**
 * Runs a report and returns the result as a Base64-encoded Excel file
 * (sync path) or queues a background job (async path) for the
 * ExportExcelButton to poll via getJobStatus().
 *
 * ## Return shape
 * Sync  → { async: false, base64: string, filename: string }
 * Async → { async: true,  jobId: string }
 *
 * ## Gap #11 fix — no duplicate ReportJob rows
 *
 * The previous implementation called `runReport()` which, on the async path,
 * already created a `ReportJob` row (inside runner.ts) and returned
 * `{ async: true, jobId }`. Then `exportReportToExcel` created a SECOND
 * `ReportJob` via `prisma.reportJob.create()`, orphaning the first.
 *
 * The fix: when `runReport` returns `{ async: true, jobId }`, that jobId IS
 * the queued job. We return it directly — no second `create()` call needed.
 * The worker (app/api/mis/worker/route.ts) processes whichever job was
 * created first; there is only ever ONE job per user action now.
 *
 * ## Access Denial
 * Called from a Client Component button — ExportExcelButton already handles
 * thrown errors by transitioning to its `error` state and showing the message.
 * So we throw a user-friendly string here (safe for that context).
 */
export async function exportReportToExcel(
  reportId: string,
  filters: unknown
): Promise<ExportExcelResponse> {
  const session = await getSession();

  // 1. Look up column definitions from the registry.
  const reportDef = REGISTRY[reportId];
  if (!reportDef) {
    throw new Error(`Report "${reportId}" not found in registry.`);
  }

  // 2. Run the report — assertReportAccess fires inside runReport.
  //    For large reports, runReport creates the ReportJob internally and
  //    returns { async: true, jobId }. We must NOT create another job.
  let result: GenerateReportResponse;
  try {
    result = await runReport(
      reportId,
      filters,
      session.orgId,
      session.userId,
      session.permissions
    );
  } catch (error: any) {
    if (error?.code === 'MIS_ACCESS_DENIED') {
      // Throw a user-readable message — ExportExcelButton shows it in-line.
      throw new Error(
        'You do not have permission to export this report. Contact your administrator.'
      );
    }
    throw new Error(error.message ?? 'Failed to export report.');
  }

  // 3. Large report — runner already queued a background Excel job.
  //    Gap #11 fix: return the jobId that runner created. No second
  //    prisma.reportJob.create() here — that was the source of the duplicate.
  if (result.async) {
    // result.jobId is guaranteed to exist when result.async === true
    // (see runner.ts — it always sets jobId when creating the ReportJob).
    return { async: true, jobId: result.jobId! };
  }

  // 4. Generate the Excel buffer (sync path — data is already in memory).
  const buffer = await generateExcelBuffer(
    reportDef.columns,
    result.rows ?? [],
    result.totals ?? {}
  );

  // 5. Build a safe, date-stamped filename.
  const dateSuffix = new Date().toISOString().split('T')[0];
  const safeName = reportDef.name.replace(/[^a-zA-Z0-9]+/g, '_');

  // 6. Return Base64 — safe for Server Action serialisation.
  return {
    async: false,
    base64: buffer.toString('base64'),
    filename: `${safeName}_${dateSuffix}.xlsx`,
  };
}

// ─── Report Columns (for Drill-Down) ─────────────────────────────────────────

/**
 * Returns the serialisable column spec and display name for a given report.
 * Used by `UniversalReportShell` to render the `DrillDownPanel` table headers
 * without requiring a second full `generateReport` call.
 *
 * ## RBAC (Gap #18 fix)
 * This function now enforces RBAC. Although column *names* are not data,
 * leaking the existence of admin-only report schemas (e.g. a `doctor_payout`
 * column on a report the caller cannot access) is an information-disclosure
 * vulnerability per PRD §4.
 *
 * If the caller's role does not include the report's `requiredPermission`,
 * this action throws so the DrillDownPanel can catch and render an
 * inline access-denied message — consistent with how generateReport() works.
 *
 * The actual *data* fetch (via `generateReport`) enforces RBAC independently,
 * so this is a defence-in-depth addition, not a primary gate.
 */
export async function getReportColumns(
  reportId: string
): Promise<{ columns: ColumnSpec[]; name: string }> {
  const reportDef = REGISTRY[reportId];
  if (!reportDef) {
    throw new Error(`Report "${reportId}" not found in registry.`);
  }

  // Gap #18 fix: enforce permission check before returning column metadata.
  const session = await getSession();
  const { assertReportAccess } = await import('@/lib/mis/rbac');
  assertReportAccess(reportDef.requiredPermission, session.permissions);


  return { columns: reportDef.columns, name: reportDef.name };
}


// ─── Presets (Phase E1) ────────────────────────────────────────────────────────

export async function getPresets(reportId?: string) {
  const session = await getSession();
  const whereClause: any = {
    organizationId: session.orgId,
    OR: [
      { user_id: session.userId },
      { is_shared: true }
    ]
  };
  if (reportId) {
    whereClause.report_id = reportId;
  }
  const presets = await prisma.reportPreset.findMany({
    where: whereClause,
    orderBy: { createdAt: 'desc' }
  });
  return presets;
}

export async function savePreset(data: { report_id: string; name: string; filters_json: any; is_shared?: boolean }) {
  const session = await getSession();
  const preset = await prisma.reportPreset.create({
    data: {
      report_id: data.report_id,
      name: data.name,
      filters_json: data.filters_json,
      is_shared: Boolean(data.is_shared),
      user_id: session.userId,
      organizationId: session.orgId
    }
  });
  return preset;
}

export async function updatePreset(id: string, data: { name?: string; is_shared?: boolean }) {
  const session = await getSession();
  const preset = await prisma.reportPreset.findUnique({ where: { id } });
  if (!preset) throw new Error('Preset not found');
  if (preset.user_id !== session.userId || preset.organizationId !== session.orgId) {
    throw new Error('Forbidden: Only the owner can modify this preset');
  }
  const updated = await prisma.reportPreset.update({
    where: { id },
    data: {
      name: data.name !== undefined ? data.name : preset.name,
      is_shared: data.is_shared !== undefined ? Boolean(data.is_shared) : preset.is_shared,
    }
  });
  return updated;
}

export async function deletePreset(id: string) {
  const session = await getSession();
  const preset = await prisma.reportPreset.findUnique({ where: { id } });
  if (!preset) throw new Error('Preset not found');
  if (preset.user_id !== session.userId || preset.organizationId !== session.orgId) {
    throw new Error('Forbidden: Only the owner can delete this preset');
  }
  await prisma.reportPreset.delete({ where: { id } });
  return { success: true };
}


// ─── Schedules (Phase E2) ──────────────────────────────────────────────────────

export async function saveSchedule(data: {
  report_id: string;
  preset_id?: string;
  filters_json?: any;
  cron_spec: string;
  format: string;
  recipients_json: any;
  channel: string;
}) {
  const session = await getSession();

  // Calculate first run
  const cronParser = require('cron-parser');
  const parseCron = cronParser.parseExpression || (cronParser.default && cronParser.default.parse) || cronParser.parse;
  const interval = parseCron(data.cron_spec, { tz: 'Asia/Kolkata' });
  const nextRunAt = interval.next().toDate();

  const schedule = await prisma.reportSchedule.create({
    data: {
      report_id: data.report_id,
      preset_id: data.preset_id || null,
      filters_json: data.filters_json || null,
      cron_spec: data.cron_spec,
      format: data.format,
      recipients_json: data.recipients_json,
      channel: data.channel,
      is_active: true,
      next_run_at: nextRunAt,
      owner_user_id: session.userId,
      organizationId: session.orgId
    }
  });
  return schedule;
}

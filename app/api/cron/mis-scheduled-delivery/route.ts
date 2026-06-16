import { NextResponse } from 'next/server';
import { prisma } from '@/backend/db';
import { runReport, REGISTRY } from '@/lib/mis/runner';
import { generateExcelBuffer } from '@/lib/mis/exporter';
import { deliverReport } from '@/lib/mis/delivery';
import { getMISPermissions } from '@/lib/mis/rbac';
import cronParser from 'cron-parser';

// This must be a GET request for Vercel Cron
export async function GET(req: Request) {
    try {
        // 1. Verify Authorization Header
        const authHeader = req.headers.get('authorization');
        const cronSecret = process.env.CRON_SECRET;
        
        if (!cronSecret || authHeader !== `Bearer ${cronSecret}`) {
            return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
        }

        // 2. Query due schedules
        const now = new Date();
        const dueSchedules = await prisma.reportSchedule.findMany({
            where: {
                is_active: true,
                OR: [
                    { next_run_at: { lte: now } },
                    { next_run_at: null }
                ]
            },
            include: {
                preset: true,
                owner_user: true
            }
        });

        if (dueSchedules.length === 0) {
            return NextResponse.json({ success: true, executed: 0 });
        }

        let successCount = 0;

        // 3. Process each schedule
        for (const schedule of dueSchedules) {
            try {
                const reportDef = REGISTRY[schedule.report_id];
                if (!reportDef) {
                    throw new Error(`Report "${schedule.report_id}" not found in registry.`);
                }

                // Determine filters: if preset_id is set, use preset's filters, else use standalone filters
                const filters = schedule.preset_id && schedule.preset 
                    ? schedule.preset.filters_json 
                    : schedule.filters_json;

                const permissions = getMISPermissions(schedule.owner_user.role);

                // Run report
                const result = await runReport(
                    schedule.report_id,
                    filters,
                    schedule.organizationId,
                    schedule.owner_user_id,
                    permissions
                );

                // Generate Excel
                const buffer = await generateExcelBuffer(
                    reportDef.columns,
                    result.rows ?? [],
                    result.totals ?? {}
                );

                // Parse recipients
                let recipients: string[] = [];
                try {
                    recipients = Array.isArray(schedule.recipients_json) 
                        ? schedule.recipients_json as string[]
                        : JSON.parse(schedule.recipients_json as string);
                } catch {
                    recipients = [];
                }

                if (recipients.length === 0) {
                    throw new Error("No valid recipients configured");
                }

                // Deliver
                await deliverReport(
                    reportDef.name,
                    buffer,
                    schedule.format,
                    recipients,
                    schedule.channel
                );

                // Compute next_run_at
                const interval = cronParser.parseExpression(schedule.cron_spec, { tz: 'Asia/Kolkata' });
                const nextRunAt = interval.next().toDate();

                // Update schedule success
                await prisma.reportSchedule.update({
                    where: { id: schedule.id },
                    data: {
                        last_run_at: new Date(),
                        last_status: 'Success',
                        run_count: schedule.run_count + 1,
                        next_run_at: nextRunAt
                    }
                });

                successCount++;

            } catch (jobError: any) {
                console.error(`[CRON ERROR] Schedule ${schedule.id} failed:`, jobError);
                
                // Update schedule failure
                // Calculate next run so it doesn't just immediately retry infinitely
                let nextRunAt = null;
                try {
                    const interval = cronParser.parseExpression(schedule.cron_spec, { tz: 'Asia/Kolkata' });
                    nextRunAt = interval.next().toDate();
                } catch (e) {
                    // Fallback to null
                }

                await prisma.reportSchedule.update({
                    where: { id: schedule.id },
                    data: {
                        last_run_at: new Date(),
                        last_status: `Failed: ${jobError.message.substring(0, 200)}`,
                        next_run_at: nextRunAt
                    }
                });
            }
        }

        return NextResponse.json({ success: true, executed: successCount });

    } catch (error: any) {
        console.error('[CRON FATAL]', error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}

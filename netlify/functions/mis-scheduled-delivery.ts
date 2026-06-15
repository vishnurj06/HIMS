import type { Config } from "@netlify/functions";

export default async (req: Request) => {
    await fetch(`${process.env.URL}/api/cron/mis-scheduled-delivery`, {
        headers: {
            Authorization: `Bearer ${process.env.CRON_SECRET}`
        }
    });
    return new Response("Scheduled Delivery Triggered Successfully");
};

export const config: Config = {
    schedule: "30 1 * * *"
};
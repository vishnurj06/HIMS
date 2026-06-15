import type { Config } from "@netlify/functions";

export default async (req: Request) => {
    await fetch(`${process.env.URL}/api/mis/worker`, {
        headers: {
            Authorization: `Bearer ${process.env.CRON_SECRET}`
        }
    });
    return new Response("Worker Triggered Successfully");
};

export const config: Config = {
    schedule: "* * * * *"
};
import { neon } from "@neondatabase/serverless";
import "dotenv/config";

export const sql = neon(process.env.DATABASE_URL); // This create sql connection

export const initDB = async () => {
    try {
        await sql`CREATE TABLE IF NOT EXISTS  transactions(
        id SERIAL PRIMARY KEY,
        user_id VARCHAR(100) NOT NULL,
        title VARCHAR(100) NOT NULL,
        amount DECIMAL(10, 2) NOT NULL,
        category VARCHAR(100) NOT NULL,
        created_at DATE NOT NULL DEFAULT CURRENT_DATE)`;

        console.log("Database initialized successfully");
    } catch (error) {
        console.error("Error initializing database:", error);
        throw error;
    }
};

import { Router } from "express";
import { sql } from "../config/db.js";

const router = Router();

router.post("/create", async (req, res) => {
    try {
        const { user_id, title, amount, category } = req.body;

        if (!user_id || !title || amount === undefined || !category) {
            return res
                .status(400)
                .json({ message: "All fields are required!" });
        }

        const transaction =
            await sql`INSERT INTO transactions(user_id, title, amount, category) VALUES (${user_id}, ${title}, ${amount}, ${category}) RETURNING *`;

        res.status(201).json(transaction[0]);
    } catch (error) {
        console.log("Error creating the transaction!", error);
        res.status(500).json({ message: "Internal server error!" });
    }
});

router.get("/:userId", async (req, res) => {
    try {
        const { userId } = req.params;

        const transactions =
            await sql`SELECT * FROM transactions WHERE user_id= ${userId} ORDER BY created_at DESC`;

        res.status(200).json(transactions);
    } catch (error) {
        console.log("Error getting the transaction!", error);
        res.status(500).json({ message: "Internal server error!" });
    }
});

router.delete("/delete/:id", async (req, res) => {
    try {
        const { id } = req.params;

        if (isNaN(parseInt(id)))
            return res.status(400).json({ message: "Invalid transaction ID!" });

        const transaction =
            await sql`DELETE FROM transactions WHERE id= ${id} RETURNING *`;

        if (transaction.length === 0)
            return res.status(404).json({ message: "Transaction not found!" });

        res.status(200).json({ message: "Transaction deleted successfully!" });
    } catch (error) {
        console.log("Error deleting the transaction!", error);
        res.status(500).json({ message: "Internal server error!" });
    }
});

router.get("/summary/:user_id", async (req, res) => {
    try {
        const { user_id } = req.params;

        const balance =
            await sql`SELECT COALESCE(SUM(amount), 0) as balance FROM transactions WHERE user_id= ${user_id}`;
        const income =
            await sql`SELECT COALESCE(SUM(amoun), 0) as income FROM transactions WHERE user_id= ${user_id} AND amount > 0`;
        const expense =
            await sql`SELECT COALESCE(SUM(amoun), 0) as expense FROM transactions WHERE user_id= ${user_id} AND amount < 0`;

        res.status(200).json({
            balance: balance[0].balance,
            income: income[0].income,
            expense: expense[0].expense,
        });
    } catch (error) {
        console.log("Error getting transaction summary!", error);
        res.status(500).json({ message: "Internal server error!" });
    }
});

export default router;

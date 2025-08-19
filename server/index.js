import cors from "cors";
import dotenv from "dotenv";
import express from "express";
import ratelimiter from "./src/middleware/rateLimiter.js";
import transaction from "./src/routes/transaction.js";
import { initDB } from "./src/config/db.js";

dotenv.config();
const app = express();

app.use(ratelimiter);
app.use(express.json());
app.use(cors({ credentials: true }));
app.use(express.json({ limit: "50mb" }));

app.use("/api/v1/transactions", transaction);

initDB().then(() => {
    app.listen(process.env.PORT, () => {
        console.log(`Server is up and running on PORT: ${process.env.PORT}`);
    });
});

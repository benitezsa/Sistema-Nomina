import { NextResponse } from "next/server";
import pool from "@/lib/db";

export async function GET() {
  try {
    const [rows] = await pool.query("SELECT DATABASE() AS database_name");

    return NextResponse.json({
      success: true,
      message: "Conexión a MySQL exitosa",
      data: rows,
    });
  } catch (error) {
    console.error("Error de conexión MySQL:", error);

    return NextResponse.json(
      {
        success: false,
        message: "No se pudo conectar a MySQL",
      },
      { status: 500 }
    );
  }
}
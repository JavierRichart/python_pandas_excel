import polars as pl

SRC = "data/raw/ventas.csv"
OUT_DETALLE = "data/processed/ventas_filtrado.csv"
OUT_RESUMEN = "data/processed/ventas_por_producto.csv"


df = pl.read_csv

df_proc = (
    df.with_columns(
        (pl.col("precio") * pl.col("ud_vendidas")).alias("importe")
    )
    .filter(pl.col("precio") >= 11)
)

resumen = (
    df_proc.group_by("producto")
    .agg([
        pl.sum("importe").alias("importe_total"),
        pl.sum("ud_vendidas").alias("uds_total")
    ])
    .sort("importe_total", descending=True)
)

df_proc.write_csv(OUT_DETALLE)
resumen.write_csv(OUT_RESUMEN)
print("Listo", OUT_DETALLE, OUT_RESUMEN, sep="\n")
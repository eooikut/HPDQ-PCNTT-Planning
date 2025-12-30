# from flask import Blueprint, render_template, request
# import pandas as pd
# from sqlalchemy import inspect
# from storage_utils import load_metadata
# from db import engine   # ✅ Dùng chung engine SQL Server

# home_bp = Blueprint("home", __name__)

# def load_general_report_from_db(table_name="report") -> pd.DataFrame:
#     """
#     Đọc toàn bộ bảng report từ SQL Server, trả về DataFrame.
#     Kiểm tra bảng tồn tại và chuẩn hóa cột 'Ngày'.
#     """
#     # ✅ Kiểm tra bảng tồn tại
#     insp = inspect(engine)
#     if table_name not in insp.get_table_names():
#         return pd.DataFrame()

#     # Đọc dữ liệu
#     df = pd.read_sql_table(table_name, con=engine)

#     # Chuyển kiểu datetime cho cột 'Ngày' nếu có
#     if "Ngày" in df.columns:
#         df["Ngày"] = pd.to_datetime(df["Ngày"], errors="coerce")
#     return df


# def load_general_report(selected_lsx="ALL") -> pd.DataFrame:
#     df_all = load_general_report_from_db("report")
#     if df_all.empty:
#         return pd.DataFrame()
#     if selected_lsx != "ALL":
#         df_all = df_all[df_all["lsx_id"] == selected_lsx]
#     return df_all


# @home_bp.route("/", methods=["GET"])
# def xem_theo_ngay():
#     metadata = load_metadata()

#     # Nếu chưa có LSX thì trả về giao diện trống
#     if not metadata:
#         return render_template(
#             "xem_theo_ngay.html",
#             has_data=False,
#             rows_by_day=[],
#             chart_datasets=[],
#             total_required=0,
#             total_actual=0,
#             selected_lsx=None,
#             start_date=None,
#             end_date=None,
#             lsx_list=[]
#         )

#     # Lấy query params
#     selected_lsx = request.args.get("lsx_id", "ALL")
#     start_date   = request.args.get("start_date")
#     end_date     = request.args.get("end_date")

#     # ✅ Đọc dữ liệu từ SQL Server
#     df_all = load_general_report_from_db("report")

#     # Xử lý dữ liệu
#     df_all = df_all.drop_duplicates(subset=["Ngày", "Order"], keep="last")

#     if selected_lsx != "ALL":
#         df_all = df_all[df_all["lsx_id"] == selected_lsx]

#     df_all["Order"] = pd.to_numeric(df_all["Order"], errors="coerce").astype("Int64")
#     df_all["Ngày"]  = pd.to_datetime(df_all["Ngày"], errors="coerce")

#     # Lọc theo ngày nếu có
#     if start_date and end_date:
#         start_dt = pd.to_datetime(start_date)
#         end_dt   = pd.to_datetime(end_date)
#         filtered = df_all[(df_all["Ngày"] >= start_dt) & (df_all["Ngày"] <= end_dt)]
#     else:
#         filtered = df_all.copy()

#     # Gom nhóm theo ngày
#     rows_by_day = []
#     if not filtered.empty:
#         for day, group in filtered.groupby(filtered["Ngày"].dt.date):
#             orders = []
#             total_loss = 0
#             for _, row in group.iterrows():
#                 sl_tb = row.get("SL trung bình/ngày", 0) or 0
#                 sl_tt = row.get("Sản lượng thực tế", 0) or 0
#                 loss  = sl_tb - sl_tt if sl_tt < sl_tb else 0
#                 total_loss += loss

#                 if pd.isna(row.get("SL trung bình/ngày")):
#                     trang_thai = "Chờ LSX mới"
#                 elif sl_tt < sl_tb:
#                     trang_thai = f"Thất thoát {loss:,.0f} kg"
#                 elif sl_tt > sl_tb:
#                     trang_thai = "Vượt tiến độ"
#                 else:
#                     trang_thai = "Đúng tiến độ"

#                 orders.append({
#                     "Order": row["Order"],
#                     "SL trung bình/ngày": sl_tb,
#                     "Sản lượng thực tế": sl_tt,
#                     "Trạng thái": trang_thai
#                 })

#             rows_by_day.append({
#                 "Ngày": day.strftime("%Y-%m-%d"),
#                 "Tổng thất thoát": total_loss,
#                 "Orders": orders
#             })

#     # Chuẩn bị dữ liệu chart
#     chart_data = {
#         r["Ngày"]: {
#             "required": sum(o["SL trung bình/ngày"] for o in r["Orders"]),
#             "actual":   sum(o["Sản lượng thực tế"] for o in r["Orders"])
#         }
#         for r in rows_by_day
#     }

#     chart_datasets = [
#         {
#             "label": "Sản lượng yêu cầu",
#             "data": [{"x": d, "y": v["required"]} for d, v in chart_data.items()],
#             "borderColor": "blue",
#             "tension": 0.3,
#             "fill": False
#         },
#         {
#             "label": "Sản lượng thực tế",
#             "data": [{"x": d, "y": v["actual"]} for d, v in chart_data.items()],
#             "borderColor": "red",
#             "tension": 0.3,
#             "fill": False
#         }
#     ]

#     total_required = sum(v["required"] for v in chart_data.values())
#     total_actual   = sum(v["actual"] for v in chart_data.values())

#     return render_template(
#         "xem_theo_ngay.html",
#         selected_lsx=selected_lsx,
#         start_date=start_date,
#         end_date=end_date,
#         lsx_list=[{"id": "ALL", "name": "ALL LSX"}] + [
#             {"id": m["id"], "name": m.get("name", m["id"])} for m in metadata
#         ],
#         rows_by_day=rows_by_day,
#         chart_datasets=chart_datasets,
#         has_data=True,
#         total_required=total_required,
#         total_actual=total_actual
#     )

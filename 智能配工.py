import streamlit as st
import pandas as pd
from io import BytesIO
import copy
import random
from collections import defaultdict

# 页面配置
st.set_page_config(page_title="桥吊理货配工系统", layout="wide")
st.title("🚢 桥吊理货智能配工系统")
st.write("请上传Excel文件，系统将自动分配理货组长和理货员，并展示人员状态")

# 工具函数
def clean_crane_name(name):
    """清洗桥吊号，保证每个编号独立"""
    name = str(name).strip()
    if not name:
        return []
    name = name.replace("，", ",").upper().strip()
    return [c.strip() for c in name.split(",") if c.strip()]

def assign_cranes_fixed(total_cranes, staff_list, min_per=4, max_per=6):
    """分配桥吊给理货员（保证每人必分配一次）"""
    n = len(staff_list)
    if total_cranes < n * min_per or total_cranes > n * max_per:
        st.error(f"桥吊数 {total_cranes} 无法满足每人 {min_per}-{max_per} 的分配")
        return None
    
    # 每人初始分配最小值
    counts = [min_per] * n
    remaining = total_cranes - sum(counts)
    
    # 循环分配剩余桥吊，确保不超过 max_per
    i = 0
    while remaining > 0:
        if counts[i] < max_per:
            counts[i] += 1
            remaining -= 1
        i = (i + 1) % n
    
    return dict(zip(staff_list, counts))

def categorize_ship_size(ship_cranes):
    """根据桥吊数量判断船舶大小：3个以上桥吊为大船，1-3个为小船"""
    crane_count = len(ship_cranes)
    if crane_count > 3:
        return "大船", crane_count
    else:
        return "小船", crane_count

# 上传Excel文件
uploaded_file = st.file_uploader("选择Excel文件（需按照规定格式）", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # 工作表名称匹配
        sheet_name_map = {
            "泊位与桥吊关联表": None,
            "船舶与桥吊关联表": None,
            "人员信息表": None,
            "四期-组长带船限制": None,
            "理货员桥吊负责规则": None
        }
        
        for key in sheet_name_map:
            matches = [s for s in xls.sheet_names if key in s]
            if matches:
                sheet_name_map[key] = matches[0]
            else:
                st.error(f"未找到工作表：{key}")
                st.stop()
        
        # 读取数据
        df_berth_crane = pd.read_excel(uploaded_file, sheet_name=sheet_name_map["泊位与桥吊关联表"])
        df_ship_crane = pd.read_excel(uploaded_file, sheet_name=sheet_name_map["船舶与桥吊关联表"])
        df_staff = pd.read_excel(uploaded_file, sheet_name=sheet_name_map["人员信息表"])
        df_leader_ship_limit = pd.read_excel(uploaded_file, sheet_name=sheet_name_map["四期-组长带船限制"])
        df_staff_crane_limit = pd.read_excel(uploaded_file, sheet_name=sheet_name_map["理货员桥吊负责规则"])
        
        # 人员状态展示
        st.subheader("📊 今日人员状态")
        状态配置 = [
            ("请假人员", "是否请假（是/否）"),
            ("公司抽调", "公司抽调（是/否）"),
            ("负责闸口", "负责闸口（是/否）"),
            ("驾驶员", "驾驶员（是/否）"),
            ("设备员", "设备员（是/否）"),
            ("申请加班", "申请加班（是/否）"),
        ]
        
        cols = st.columns(6)
        for i, (label, col) in enumerate(状态配置):
            with cols[i]:
                people = df_staff[df_staff[col] == "是"]["姓名"].tolist()
                st.write(f"**{label}**（{len(people)}人）")
                st.write(", ".join(people) if people else "无")
        
        # 筛选可用人员
        # 理货组长
        leader_available = df_staff[
            (df_staff["岗位类型"] == "理货组长") &
            (df_staff["是否请假（是/否）"] == "否") &
            (df_staff["公司抽调（是/否）"] == "否") &
            (df_staff["负责闸口（是/否）"] == "否")
        ].groupby("工作地（四期/自动化/闸口）")["姓名"].apply(list).to_dict()
        
        # 理货员（原始列表和可用列表）
        staff_original = df_staff[
            (df_staff["岗位类型"] == "理货员") &
            (df_staff["是否请假（是/否）"] == "否") &
            (df_staff["公司抽调（是/否）"] == "否") &
            (df_staff["负责闸口（是/否）"] == "否")
        ].groupby("工作地（四期/自动化/闸口）")["姓名"].apply(list).to_dict()
        
        staff_available = copy.deepcopy(staff_original)
        
        st.subheader("👥 可用配工人员")
        col1, col2 = st.columns(2)
        with col1:
            st.write("**理货组长可用数量**")
            for wa in ["四期", "自动化"]:
                names = leader_available.get(wa, [])
                st.write(f"{wa}：{len(names)}人（{', '.join(names) if names else '无'}）")
        
        with col2:
            st.write("**理货员可用数量（初始）**")
            for wa in ["四期", "自动化"]:
                names = staff_original.get(wa, [])
                st.write(f"{wa}：{len(names)}人（{', '.join(names) if names else '无'}）")
        
        # 船舶与桥吊关联处理
        all_cranes = {}
        for _, row in df_berth_crane.iterrows():
            workarea = str(row["工作地"]).strip()
            raw_cranes = str(row["桥吊号（按从左到右顺序，逗号分隔）"])
            cranes = clean_crane_name(raw_cranes)
            for c in cranes:
                all_cranes[c] = workarea
        
        # 读取船舶表并清洗桥吊号，建立船舶-桥吊映射关系，新增船舶大小分类
        ship_crane_list = []
        ship_to_cranes = {}  # 船舶到桥吊的映射
        ship_size_info = {}  # 船舶大小信息
        for _, row in df_ship_crane.iterrows():
            ship_name = str(row["船舶名称"]).strip()
            raw = str(row["对应桥吊号（逗号分隔，需属于工作表1中的桥吊）"])
            cranes = clean_crane_name(raw)
            
            # 判断船舶大小
            size, crane_count = categorize_ship_size(cranes)
            ship_size_info[ship_name] = {"size": size, "crane_count": crane_count}
            
            matched_workarea = None
            matched_cranes = []
            for c in cranes:
                if c in all_cranes and not matched_workarea:
                    matched_workarea = all_cranes[c]
                matched_cranes.append(c)
            
            if matched_workarea:
                ship_crane_list.append({
                    "船舶名称": ship_name,
                    "桥吊列表": matched_cranes,
                    "工作地": matched_workarea,
                    "大小": size,
                    "桥吊数量": crane_count
                })
                ship_to_cranes[ship_name] = matched_cranes  # 存储船舶对应的桥吊
        
        # 按工作地分组船舶和桥吊，新增船舶大小统计
        workarea_data = {
            "四期": {
                "ships": [],
                "all_cranes": [],
                "crane_to_ship": {},  # 桥吊到船舶的映射
                "large_ships": [],    # 大船列表
                "small_ships": []     # 小船列表
            },
            "自动化": {
                "ships": [],
                "all_cranes": [],
                "crane_to_ship": {},  # 桥吊到船舶的映射
                "large_ships": [],    # 大船列表
                "small_ships": []     # 小船列表
            }
        }
        
        for s in ship_crane_list:
            wa = s["工作地"]
            if wa in workarea_data:
                workarea_data[wa]["ships"].append(s)
                # 按大小分类船舶
                if s["大小"] == "大船":
                    workarea_data[wa]["large_ships"].append(s["船舶名称"])
                else:
                    workarea_data[wa]["small_ships"].append(s["船舶名称"])
                    
                for c in s["桥吊列表"]:
                    if c not in workarea_data[wa]["all_cranes"]:
                        workarea_data[wa]["all_cranes"].append(c)
                    # 记录桥吊对应的船舶
                    if c not in workarea_data[wa]["crane_to_ship"]:
                        workarea_data[wa]["crane_to_ship"][c] = []
                    workarea_data[wa]["crane_to_ship"][c].append(s["船舶名称"])
        
        st.subheader("🚢 各工作地待配工数据")
        for wa in ["四期", "自动化"]:
            ships = workarea_data[wa]["ships"]
            cranes = workarea_data[wa]["all_cranes"]
            large = len(workarea_data[wa]["large_ships"])
            small = len(workarea_data[wa]["small_ships"])
            st.write(f"{wa}：{len(ships)}艘船舶（大船{large}艘/小船{small}艘），{len(cranes)}个桥吊（{', '.join(cranes[:100])}...）")
        
        # 配工逻辑
        def assign_work(workarea, staff_available):
            data = workarea_data[workarea]
            ships = data["ships"]
            all_cranes = data["all_cranes"]
            crane_to_ship = data["crane_to_ship"]  # 桥吊到船舶的映射
            total_cranes = len(all_cranes)
            total_ships = len(ships)
            
            leaders = leader_available.get(workarea, [])
            current_staff = staff_available.get(workarea, [])
            num_leaders = len(leaders)
            
            if not ships:
                st.warning(f"{workarea} 无待配工船舶")
                return None, staff_available
            if not leaders:
                st.warning(f"{workarea} 无可用理货组长")
                return None, staff_available
            if not current_staff:
                st.warning(f"{workarea} 无可用理货员")
                return None, staff_available
            
            # 计算每个组长应分配的平均船舶数量
            avg_ships_per_leader = total_ships / num_leaders if num_leaders > 0 else 0
            min_ships = int(avg_ships_per_leader)
            max_ships = min_ships + 1 if total_ships % num_leaders != 0 else min_ships
            # st.write(f"{workarea} 配工策略：{num_leaders}位组长，共{total_ships}艘船，平均每人分配{min_ships}-{max_ships}艘船（兼顾大小搭配）")
            
            # 四期分配逻辑
            if workarea == "四期":
                st.write(f"四期总桥吊数：{total_cranes}个")
                staff_crane_map = assign_cranes_fixed(total_cranes, current_staff, min_per=4, max_per=6)
                
                if not staff_crane_map:
                    return None, staff_available
                
                # 更新可用理货员
                staff_available[workarea] = []
                
                # 拆分桥吊列表并记录每个理货员负责的桥吊
                cranes_flat = all_cranes.copy()
                idx = 0
                for staff, count in staff_crane_map.items():
                    staff_crane_map[staff] = cranes_flat[idx: idx + count]
                    idx += count
                
                st.success(f"四期桥吊分配完成：{len(staff_crane_map)}人，桥吊总数{total_cranes}")
                
                # 船舶分配优化：均衡数量+大小搭配
                # 1. 准备船舶数据（带大小标记）
                all_ships_with_size = [(s["船舶名称"], s["大小"], s["桥吊列表"]) for s in data["ships"]]
                random.shuffle(all_ships_with_size)  # 随机打乱顺序，增加分配均衡性
                
                # 2. 初始化组长分配池
                leader_allocations = {leader: {"ships": [], "large_count": 0, "small_count": 0, "cranes": []} 
                                     for leader in leaders}
                
                # 3. 先分配大船，确保每个组长至少有部分大船
                large_ships = [s for s in all_ships_with_size if s[1] == "大船"]
                small_ships = [s for s in all_ships_with_size if s[1] == "小船"]
                
                # 计算应分配的大船数量
                total_large = len(large_ships)
                avg_large_per_leader = total_large / num_leaders if num_leaders > 0 else 0
                min_large = int(avg_large_per_leader)
                
                # 分配大船
                current_leader_idx = 0
                for ship, size, cranes in large_ships:
                    # 找到当前船舶最少的组长
                    sorted_leaders = sorted(leaders, key=lambda x: len(leader_allocations[x]["ships"]))
                    
                    # 优先分配给大船数量较少的组长
                    for leader in sorted_leaders:
                        if leader_allocations[leader]["large_count"] < min_large + 1:
                            leader_allocations[leader]["ships"].append(ship)
                            leader_allocations[leader]["large_count"] += 1
                            leader_allocations[leader]["cranes"].extend(cranes)
                            break
                    current_leader_idx = (current_leader_idx + 1) % num_leaders
                
                # 4. 分配小船，平衡总数量并补充大小搭配
                current_leader_idx = 0
                for ship, size, cranes in small_ships:
                    # 找到当前船舶总数最少的组长
                    sorted_leaders = sorted(leaders, key=lambda x: len(leader_allocations[x]["ships"]))
                    
                    # 优先分配给小船数量较少的组长
                    for leader in sorted_leaders:
                        if len(leader_allocations[leader]["ships"]) < max_ships:
                            leader_allocations[leader]["ships"].append(ship)
                            leader_allocations[leader]["small_count"] += 1
                            leader_allocations[leader]["cranes"].extend(cranes)
                            break
                    current_leader_idx = (current_leader_idx + 1) % num_leaders
                
                # 5. 整理组长分配结果（去重桥吊）
                leader_ship_map = {}
                for leader, alloc in leader_allocations.items():
                    unique_cranes = list(set(alloc["cranes"]))  # 去重桥吊
                    leader_ship_map[leader] = {
                        "ships": alloc["ships"],
                        "cranes": unique_cranes,
                        "large_count": alloc["large_count"],
                        "small_count": alloc["small_count"]
                    }
                
                # 展示组长分配情况
                st.write("### 组长分配详情")
                allocation_details = []
                for leader, alloc in leader_ship_map.items():
                    allocation_details.append({
                        "理货组长": leader,
                        "总船舶数": len(alloc["ships"]),
                        "大船数": alloc["large_count"],
                        "小船数": alloc["small_count"],
                        "负责桥吊数": len(alloc["cranes"])
                    })
                st.dataframe(pd.DataFrame(allocation_details), use_container_width=True)
                
                # 6. 整理最终配工结果
                final_result = []
                for staff, cranes in staff_crane_map.items():
                    assigned_leader = "未分配组长"
                    assigned_ships = []
                    
                    # 匹配理货员桥吊对应的组长
                    for leader, group in leader_ship_map.items():
                        if any(c in group["cranes"] for c in cranes):
                            assigned_leader = leader
                            assigned_ships = group["ships"]
                            break
                    
                    final_result.append({
                        "工作地": workarea,
                        "理货组长": assigned_leader,
                        "负责船舶": ", ".join(assigned_ships),
                        # "船舶数量": len(assigned_ships),
                        "理货员": staff,
                        "负责桥吊": ", ".join(cranes),
                        "桥吊数量": len(cranes)
                    })
                
                return pd.DataFrame(final_result), staff_available
            
            # 自动化分配逻辑
            else:
                st.write(f"自动化总桥吊数：{total_cranes}个")
                base_qty = 2
                
                # 校验桥吊数是否为2的倍数
                if total_cranes % base_qty != 0:
                    st.error(f"自动化桥吊总数{total_cranes}个，需为2的倍数")
                    return None, staff_available
                
                # 校验理货员数量是否足够
                num_staff_needed = total_cranes // base_qty
                if num_staff_needed > len(current_staff):
                    st.error(f"自动化理货员不足（需{num_staff_needed}人，仅{len(current_staff)}人）")
                    return None, staff_available
                
                # 分配理货员及桥吊
                assigned_staff = current_staff[:num_staff_needed]
                staff_available[workarea] = current_staff[num_staff_needed:]  # 更新剩余可用理货员
                
                # 拆分桥吊给理货员（每人2个）
                staff_crane_map = {}
                crane_idx = 0
                for staff in assigned_staff:
                    staff_crane_map[staff] = all_cranes[crane_idx:crane_idx + base_qty]
                    crane_idx += base_qty
                
                st.success(f"自动化桥吊分配完成：{len(staff_crane_map)}人，桥吊总数{total_cranes}")
                
                # 自动化船舶分配（均衡数量+大小搭配）
                # 1. 准备船舶数据
                all_ships_with_size = [(s["船舶名称"], s["大小"], s["桥吊列表"]) for s in data["ships"]]
                random.shuffle(all_ships_with_size)
                
                # 2. 初始化组长分配池
                leader_allocations = {leader: {"ships": [], "large_count": 0, "small_count": 0, "cranes": []} 
                                     for leader in leaders}
                
                # 3. 分离大小船
                large_ships = [s for s in all_ships_with_size if s[1] == "大船"]
                small_ships = [s for s in all_ships_with_size if s[1] == "小船"]
                
                # 4. 分配大船
                total_large = len(large_ships)
                avg_large_per_leader = total_large / num_leaders if num_leaders > 0 else 0
                min_large = int(avg_large_per_leader)
                
                current_leader_idx = 0
                for ship, size, cranes in large_ships:
                    sorted_leaders = sorted(leaders, key=lambda x: len(leader_allocations[x]["ships"]))
                    for leader in sorted_leaders:
                        if leader_allocations[leader]["large_count"] < min_large + 1:
                            leader_allocations[leader]["ships"].append(ship)
                            leader_allocations[leader]["large_count"] += 1
                            leader_allocations[leader]["cranes"].extend(cranes)
                            break
                    current_leader_idx = (current_leader_idx + 1) % num_leaders
                
                # 5. 分配小船
                current_leader_idx = 0
                for ship, size, cranes in small_ships:
                    sorted_leaders = sorted(leaders, key=lambda x: len(leader_allocations[x]["ships"]))
                    for leader in sorted_leaders:
                        if len(leader_allocations[leader]["ships"]) < max_ships:
                            leader_allocations[leader]["ships"].append(ship)
                            leader_allocations[leader]["small_count"] += 1
                            leader_allocations[leader]["cranes"].extend(cranes)
                            break
                    current_leader_idx = (current_leader_idx + 1) % num_leaders
                
                # 6. 整理组长分配结果
                leader_ship_map = {}
                for leader, alloc in leader_allocations.items():
                    unique_cranes = list(set(alloc["cranes"]))
                    leader_ship_map[leader] = {
                        "ships": alloc["ships"],
                        "cranes": unique_cranes,
                        "large_count": alloc["large_count"],
                        "small_count": alloc["small_count"]
                    }
                
                # 展示组长分配情况
                st.write("### 组长分配详情")
                allocation_details = []
                for leader, alloc in leader_ship_map.items():
                    allocation_details.append({
                        "理货组长": leader,
                        "总船舶数": len(alloc["ships"]),
                        "大船数": alloc["large_count"],
                        "小船数": alloc["small_count"],
                        "负责桥吊数": len(alloc["cranes"])
                    })
                st.dataframe(pd.DataFrame(allocation_details), use_container_width=True)
                
                # 7. 整理自动化配工结果
                final_result = []
                for staff, cranes in staff_crane_map.items():
                    assigned_leader = "未分配组长"
                    assigned_ships = []
                    
                    # 匹配理货员对应的组长
                    for leader, group in leader_ship_map.items():
                        if any(c in group["cranes"] for c in cranes):
                            assigned_leader = leader
                            assigned_ships = group["ships"]
                            break
                    
                    final_result.append({
                        "工作地": workarea,
                        "理货组长": assigned_leader,
                        "负责船舶": ", ".join(assigned_ships),
                        # "船舶数量": len(assigned_ships),
                        "理货员": staff,
                        "负责桥吊": ", ".join(cranes),
                        "桥吊数量": len(cranes)
                    })
                
                return pd.DataFrame(final_result), staff_available
        
        # 执行分配并展示/下载
        if st.button("开始配工"):
            st.subheader("🔍 配工过程提示")
            df_4期, staff_available = assign_work("四期", staff_available)
            df_自动化, staff_available = assign_work("自动化", staff_available)
            
            st.subheader("🚀 四期配工结果")
            if df_4期 is not None and not df_4期.empty:
                st.dataframe(df_4期, use_container_width=True)
            else:
                st.info("四期未生成配工结果（详见上方提示）")
            
            st.subheader("🚀 自动化配工结果")
            if df_自动化 is not None and not df_自动化.empty:
                st.dataframe(df_自动化, use_container_width=True)
            else:
                st.info("自动化未生成配工结果（详见上方提示）")
            
            # 下载功能
            if (df_4期 is not None and not df_4期.empty) or (df_自动化 is not None and not df_自动化.empty):
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    if df_4期 is not None and not df_4期.empty:
                        df_4期.to_excel(writer, sheet_name="四期配工结果", index=False)
                    if df_自动化 is not None and not df_自动化.empty:
                        df_自动化.to_excel(writer, sheet_name="自动化配工结果", index=False)
                
                output.seek(0)
                st.download_button(
                    label="下载配工结果（Excel）",
                    data=output.getvalue(),
                    file_name="桥吊理货配工结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    except Exception as e:
        st.error(f"程序出错：{str(e)}")
        st.write("请检查Excel格式及数据是否符合要求")
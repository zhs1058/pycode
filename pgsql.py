import sys
import psycopg2
from psycopg2 import extensions

# 连接到数据库
try:
    conn = psycopg2.connect(
        host="120.53.224.82",
        port=5432,
        database="ibank",
        user="ibank",
        password="UHNiY0NvcnB3bEAyMDIx1"
    )
except Exception as e:
    print(str(e).replace("\n", ""))

# 设置连接的事务隔离级别为自动提交
conn.set_isolation_level(extensions.ISOLATION_LEVEL_AUTOCOMMIT)

# 创建一个游标对象
cur = conn.cursor()


# 获取表的详细结构信息
def get_table_structure(table_name):
    # 查询列信息
    cur.execute(
        f"""
        select 
            col.ordinal_position as serial, 
            col.table_name as tableName, 
            col.column_name as columnName,
            col.udt_name as dataType, 
            COALESCE(col.character_maximum_length, col.numeric_precision, col.datetime_precision) as len,
            col.numeric_scale as acc,
            col.is_nullable as isNull,
            col.column_default as defaultValue,
            des.description as desc,
            def.def_forkey as foreignkey
        from
	        information_schema.columns col 
	        left join pg_description des on col.table_name::regclass = des.objoid and col.ordinal_position = des.objsubid
	        left join 
		    (SELECT
			    kcu.column_name as def_column, 
			    ccu.table_name || '.' || ccu.column_name AS def_forkey
		    FROM 
                information_schema.table_constraints AS tc 
                JOIN information_schema.key_column_usage AS kcu ON tc.constraint_name = kcu.constraint_name
                JOIN information_schema.constraint_column_usage AS ccu ON ccu.constraint_name = tc.constraint_name
		    WHERE constraint_type = 'FOREIGN KEY' AND tc.table_name = '{table_name}') def on def.def_column = col.column_name
        where
	        table_name = '{table_name}'
        order by 
	        ordinal_position;""")
    columns = cur.fetchall()

    # 查询主键信息
    # cur.execute(
    #     f"SELECT constraint_name FROM information_schema.table_constraints WHERE table_name = '{table_name}' AND constraint_type = 'PRIMARY KEY'")
    # primary_key = cur.fetchone()
    # primary_key_name = primary_key[0] if primary_key else ""
    #
    # # 查询外键信息
    # cur.execute(f"""
    #         SELECT conname, conrelid::regclass AS table_name, a.attname AS column_name,
    #             confrelid::regclass AS foreign_table_name, af.attname AS foreign_column_name
    #         FROM pg_constraint AS c
    #         JOIN pg_attribute AS a ON a.attnum = ANY(c.conkey) AND a.attrelid = c.conrelid
    #         JOIN pg_attribute AS af ON af.attnum = ANY(c.confkey) AND af.attrelid = c.confrelid
    #         WHERE conrelid = '{table_name}'::regclass
    #         AND confrelid IS NOT NULL
    #     """)
    # foreign_keys = cur.fetchall()

    # 打印表的详细结构信息
    print(f"表名: {table_name}")
    print("序号\t表名\t列名\t数据类型\t长度\t精度\t是否为空\t默认值\t描述\t外键")
    for column in columns:

        serial, tableName, columnName, dataType, len, acc, isNull, defaultValue, desc, foreignkey = column

        # is_nullable = "是" if is_nullable == "YES" else "否"
        # is_primary_key = "是" if column_name == primary_key_name else "否"
        # foreign_key_info = [fk for fk in foreign_keys if fk[2] == column_name]
        # foreign_key_str = ", ".join([f"{fk[0]} ({fk[3]}.{fk[4]})" for fk in foreign_key_info])
        print(
            f"{serial}\t{tableName}\t{columnName}\t{dataType}\t{len}\t{acc}\t{isNull}\t{defaultValue}\t{desc}\t{foreignkey}")


# 调用函数获取表结构信息
get_table_structure("wx_seal_apply")

# 关闭游标和连接
cur.close()
conn.close()

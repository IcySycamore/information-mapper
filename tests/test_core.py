from information_mapper.core import map_records


def test_map_records_uses_first_available_alias():
    records = [{"户主姓名": "张三", "电话": "13800000000"}]
    result = map_records(
        records,
        ["姓名", "联系电话", "家庭住址"],
        {
            "姓名": ["姓名", "户主姓名"],
            "联系电话": ["联系电话", "电话"],
            "家庭住址": ["家庭住址", "地址"],
        },
    )

    assert result == [{"姓名": "张三", "联系电话": "13800000000", "家庭住址": ""}]

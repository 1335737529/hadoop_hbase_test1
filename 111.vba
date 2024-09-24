        With oSort
            '删除原来的排序字段
            .SortFields.Clear
            '添加新的排序字段,第一个关键字
            .SortFields.Add Key:=oWK.Range("a1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortTextAsNumbers
            '添加新的排序字段,第二个关键字
            .SortFields.Add Key:=oWK.Range("b1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            '设置排序的单元格区域
            .SetRange oWK.Range("a1:b" & iRow)
            '设置排序的方法
            .SortMethod = xlPinYin
            '设置排序的方向
            .Orientation = xlSortColumns
            '设置是否包含列标题
            .Header = xlYes
            '开始排序
            .Apply
        End With
    End With

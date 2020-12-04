# XlswriterExcel
高性能XlswriterExcel拓展用于excel导出，导出时有效减少内存开销


    $excel_data = [
        ['编号', '姓名', '年龄'],
        ['1', 'zhangsan', '25'],
        ['2', 'lisi', '28'],
    ];
    #实例化对象
    $XlswriterExcel = new \echoyii\XlswriterExcel();
    
    #设置文件名
    $XlswriterExcel->setBase('用户来源_'.date('YmdHis', time()));
    
    #创建工作表
    $XlswriterExcel->createSheet($excel_data, '用户来源');
    
    #输出文件
    $XlswriterExcel->fileOutput();
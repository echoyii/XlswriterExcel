<?php
/**
 * File: test.php
 * Functionality:
 * Author: Li
 * Date: 2020/4/12
 */
require_once '../vendor/autoload.php';

use echoyii\XlswriterExcel;

$excel_data = array();
//顶部
$excel_data[0] = [
    '编号',
    '类型',
    '真实姓名',
    '手机号码',
    '开户行',
    '卡号',
];

$excel_data[] = [
    "1",
    "vip会员",
    "zhansan",
    "13666666666",
    "东莞农商",
    "665146155446121100"
];

    #实例化对象
    $XlswriterExcel = new XlswriterExcel();

    #设置文件名
    $XlswriterExcel->setBase('全部银行卡列表_' . date('YmdHis', time()));

    #创建工作表
    $XlswriterExcel->createSheet($excel_data, '全部银行卡列表');

    #输出文件
    $XlswriterExcel->fileOutput();


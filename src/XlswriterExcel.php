<?php
/*
 * 功能：生成Excel文件类
 * 依赖：php_xlswriter拓展
 */

namespace echoyii;
class XlswriterExcel
{
    private $excelObject;

    private $fileName;                  //文件名称

    private $sheetList = array();       //工作表数据集合

    private $firststyle;                //顶行样式



    public function __construct()
    {


    }

    /**
     * 设置基本信息
     * @param string $filePath  文件路径
     * @param string $fileName  文件名称
     */
    public function setBase($fileName){
        #设置全局文件名(含xlsx后缀)
        $this->fileName  = $fileName;

        return $this;

    }

    /**
     * 实例化对象
     * @param string $firstSheet 第一个工作表名称
     */
    private function initExcelObject($firstSheet){
        $config = ['path' => './'];

        $Excel = new \Vtiful\Kernel\Excel($config);

        #初始化excelObject, 定义文件名和第一个工作表名
        $this->excelObject = $Excel->fileName($this->fileName, $firstSheet);

        return true;
    }

    /**
     * 实例化样式资源对象
     */
    private function initFormatObject(){
        $fileHandle = $this->excelObject->getHandle();

        #初始化formatObject,创建样式资源
        $this->formatObject = new \Vtiful\Kernel\Format($fileHandle);

        //只渲染顶行样式（不渲染其他行样式，会导致导出太慢）
        $this->firststyle = $this->formatObject->bold()->toResource();

        return true;
    }


    /**
     * 添加一个工作表
     * @param string $data      内容（二维度数组）
     * @param string $sheetName 工作表名称
     */
    public function createSheet($data = array(), $sheetName = ''){
        #检查顶行内容长度和主体内容长度是否一致（待补充）

        #取出第一个元素作为顶部数据
        $header = array_shift($data);

        #统计已添加工作表数量
        $sheetCount = count($this->sheetList);

        #设置默认工作表名
        $sheetName = $sheetName === '' ? 'sheet'.++$sheetCount : $sheetName;


        if ($sheetCount == 0){
            //初次创建工作表

            #1.实例化Excel对象
            $this->initExcelObject($sheetName);

            #2.实例化Format对象，用于创建样式资源
            $this->initFormatObject();

            #3.创建工作表
            $this->excelObject->header($header)->data($data);


        }else{
            //非初次创建工作表

            #追加工作表
            $this->excelObject->addSheet($sheetName)->header($header)->data($data);
        }

        #渲染样式（只渲染顶行样式，不渲染其他行样式，会导致导出太慢）
        $this->excelObject->setRow('A1', 12, $this->firststyle);


        #把新的工作表放入sheetList
        $this->sheetList[] = [
            'header'    => $header,     //顶行内容
            'data'      => $data,       //主体内容
            'sheetName' => $sheetName,  //工作表名称
        ];

        return $this;
    }


    /**
     * 输出文件
     * @param string $fileName 文件名
     * @param string $filePath 文件路径
     */
    public function fileOutput(){

        #调用输出文件
        $filePath = $this->excelObject->output();

        $encoded_filename = urlencode($this->fileName);
        $encoded_filename = str_replace("+", "%20", $encoded_filename);
        #设置header
        $ua = $_SERVER["HTTP_USER_AGENT"];
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        if (preg_match("/MSIE/", $ua)) {
            header("Content-Disposition: attachment; filename=\"{$encoded_filename}.xls\"");
        } elseif (preg_match("/Firefox/", $ua)) {
            header("Content-Disposition: attachment; filename*=\"utf8''{$this->fileName}.xls\"");
        } else {
            header("Content-Disposition: attachment; filename=\"{$this->fileName}.xls\"");
        }
        header('Cache-Control: max-age=0');

        #输入字节流到文件
        if (copy($filePath, 'php://output') === false) {
            // Throw exception 抛出异常（待补充）
        }

        #删除文件
        @unlink($filePath);
    }

}

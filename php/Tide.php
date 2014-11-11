<?php
// Tide.php
header("Content-Type: text/html; charset=UTF-8");
//ライブラリ読み込み
set_include_path(get_include_path().PATH_SEPARATOR.$_SERVER["DOCUMENT_ROOT"].'/tide/php/Classes/');
include_once('PHPExcel.php');
include_once('PHPExcel/IOFactory.php');

// Class form->date
class Form2date {

    public $startYear;
    public $startMonth;
    public $monLength;
    public $portid;

    public function Form2date() {
       $this->startYear = date("Y");
       $this->startMonth = date("m");
       $this->portid = "67";
    }

    public function makeDate() {
       $start = explode("/",$_POST["StartMonth"]);
       $this->startYear = $start[0];
       $this->startMonth = $start[1];
       $this->monLength = $_POST['EndMonth'];
       $this->portid = $_POST['Portid'];
    }

    public function exportDate() {
        $date = array(
            "portid" => $this->portid,
            "monthLength" => $this->monLength,
            "syear" => $this->startYear,
            "smonth" => $this->startMonth,
        );
        return $date;
    }

    public function getPortid() {
        return $this->portid;
    }

}

// Class date->xml
// 期間の範囲から
class Date2xml {
    // サイトから手に入れた日付ごとのxmlデータを一括管理
    //
    // $xml = array(
    //       "2014" => array(
    //              "11" => array(
    //                          "1" => サイトのデータ,
    //                          "2" => サイトのデータ,
    //                          ...
    //                          "30" => サイトのデータ
    //                          ),
    //              "12" => array(
    //                          "1" => サイトのデータ,
    //                          ...
    //                          "31" => サイトのデータ
    //                          )
    //              )
    //        );
    public $xml;
    public $base_url;
    public $date;


    public function Date2xml($url_) {
        $this->url = $url_;
    }

    public function setDate($inputDate = array()) {
        $this->date = $inputDate;
    }

    public function setBaseUrl($url_) {
        $this->base_url = $url_;
    }

    // 日付を整形。指定された月の日付を計算し、URI作成
    public function getXML() {
        $selectYear  = intval($this->date['syear']);
        $selectMonth = intval($this->date['smonth']);
        $forLength = intval($this->date['monthLength']);
        $portid = $this->date['portid'];

        // 初期配列の生成
        $this->xml = array();
        $this->xml += array("$selectYear" => array());

        // 年と月についてのfor文、指定回数まで増加
        for($i = 0;$i < $forLength;$i++) {
            $thisMonth = $selectMonth + $i;
            $thisYear = $selectYear;

            // 指定する月が12を超えたら...年が変わったら...
            if($thisMonth > 12) {
                $thisMonth = $thisMonth - 12;
                $thisYear++;

                // 注目する年が変わったら配列を追加する
                if(!array_key_exists($thisYear, $this->xml)){
                    $this->xml += array("$thisYear" => array());
                }
            }

            // 指定する月を追加
            $this->xml["$thisYear"] += array("$thisMonth" => array());

            // 当該月の最後の日付を求める
            $selectY_M = $thisYear . "-" .$thisMonth;
            $lastDate = intval($this->getLastDay($selectY_M));
//            echo $thisYear."-".$thisMonth."<br />";  // for debug

            // 日付についてのfor文
            //for($j = 1;$j <= $lastDate;$j++) {
            for($j = 1;$j <= $lastDate;$j++) {
//                echo $j."<br/>";      // for debug
                // portid=67&year=2009&month=12&day=01
                $thisURL = $this->url;
                $thisURL = $thisURL . 'portid=' . $portid . '&year=' . $thisYear .'&month=' . $thisMonth . '&day=' . $j;
//                echo $thisURL."<br />";     //for debug
                $this->xml["$thisYear"]["$thisMonth"] += array("$j" => $this->addXML($thisURL));
            }

        }

        return $this->xml;
    }

    // 実際にサイトから値をxmlに追加するメソッド
    private function addXML($url) {
        $xml_data = "";
        $cp = curl_init();
        curl_setopt($cp, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt( $cp, CURLOPT_HEADER, false );
        curl_setopt($cp, CURLOPT_URL, $url);
        curl_setopt($cp, CURLOPT_TIMEOUT, 60);
        $xml_data = curl_exec($cp);
        curl_close($cp);
        $original_xml = simplexml_load_string($xml_data);

        $data = get_object_vars($original_xml);
        return $data;
    }

    private function getLastDay($select){
        $lastDate = explode("-",date('Y-m-d', strtotime('last day of ' . $select)));
        // 選択した月の最後の日付
        return $lastDate[2];
    }

}

// ==============================================================================================================================
// 実行
// 選択された月のデータを得て、excelに出力
if ($_SERVER["REQUEST_METHOD"] == "POST") {

    // POSTで送信された年、月と継続期間のデータを一つの配列に整形
    $form2date = new Form2date();
    $form2date->makeDate();

    // 手にした月の配列から日付を計算し、各日についてxmlデータを取得
    $url = "http://fishing-community.appspot.com/tidexml/index?";
    $date2xml = new Date2xml($url);
    $date2xml->setDate($form2date->exportDate());
    $date2xml->setBaseUrl($url);
    $xml = $date2xml->getXML();

    // xml->excel
    // 手にした潮見xmlデータをexcelのフォーマットに従うように変換
    $language = array(
        "1" => "1月 January",
        "2" => "2月 February",
        "3" => "3月 March",
        "4" => "4月 April",
        "5" => "5月 May",
        "6" => "6月 June",
        "7" => "7月 July",
        "8" => "8月 August",
        "9" => "9月 September",
        "10" => "10月 October",
        "11" => "11月 November",
        "12" => "12月 December");

    $tide_color = array(
        "大潮" => "FF7F00",
        "中潮" => "9AFF9A",
        "小潮" => "CAE1FF",
        "若潮" => "1C86EE",
        "長潮" => "1C86EE"
    );

    $portname = array(
        "67" => "oarai"
    );

    // エクセルファイルを開く
    $book = new PHPExcel();
    $i = 0;
    foreach($xml as $ykey => $year) {
        foreach($year as $mkey => $month) {

            // sheet変更の切れ目
            // sheetをactiveにして日、曜日、潮、満潮、干潮までは記述する
            //シート設定
            $book->createSheet($i);
            $book->setActiveSheetIndex($i);
            $sheet = $book->getActiveSheet();
            $sheet->setTitle($ykey."-".$mkey);
            //$book->setActiveSheetIndex($i);
            //$sheet = $book->getActiveSheet();
            //$sheet->setTitle($ykey."-".$mkey);    //シート名指定

            // 〜年　潮見表
            $sheet->mergeCells('A1:G4');
            $sheet->setCellValue("A1", $ykey."年  潮見表");

            // 〜月　英語表記
            $sheet->mergeCells('A5:G5');
            $sheet->setCellValue("A5", $language[$mkey]);
            $sheet->getStyle( 'A5:G5' )->getFill()->setFillType( PHPExcel_Style_Fill::FILL_SOLID )->getStartColor()->setRGB( "FFEC8B");


            // ヘッダー設定 日、曜日、潮、満潮、干潮
            $sheet->setCellValue("A6","日");
            $sheet->getColumnDimension('A')->setWidth(3.29);
            $sheet->setCellValue("B6","曜日");
            $sheet->getColumnDimension('B')->setWidth(5.57);
            $sheet->setCellValue("C6","潮");
            $sheet->getColumnDimension('C')->setWidth(5.57);
            $sheet->mergeCells('D6:E6');
            $sheet->setCellValue("D6","満潮");
            $sheet->getColumnDimension('D')->setWidth(13);
            $sheet->getColumnDimension('E')->setWidth(13);
            $sheet->mergeCells('F6:G6');
            $sheet->setCellValue("F6","干潮");
            $sheet->getColumnDimension('F')->setWidth(11.86);
            $sheet->getColumnDimension('G')->setWidth(11.86);


            // $day_countによって、日付の増加とともにエクセルの行の位置を増加
            $day_count = 7;
            foreach($month as $dkey => $day) {
                $sheet->setCellValue("A$day_count",$day['day']);
                $sheet->setCellValue("B$day_count",$day['youbi']);
                $sheet->setCellValue("C$day_count",$day['tide-name']);
                // tide-nameで分岐して色を変える
                $tide_color_byname = $tide_color[$day['tide-name']];
                $sheet->getStyle( "A$day_count:G$day_count" )->getFill()->setFillType( PHPExcel_Style_Fill::FILL_SOLID )->getStartColor()->setRGB( "$tide_color_byname");

                // 始めが満潮か干潮か見分ける
                if(intval($day['tidedetails'][0]->{'tide-level'}) > (intval($day['tidedetails'][1]->{'tide-level'}))){

                    // 始めが満潮の時のセルの表示位置を決定
                    $high = array("D","F","E","G");
                } else {

                    // 始めが干潮の時のセルの表示位置を決定
                    $high = array("F","D","G","E");
                }

                // 何番目の潮かをカウント
                $highCount = 0;
                foreach($day['tidedetails'] as $val){
                    if(!empty($val->{'tide-time'})) {
//                        echo $val->{'tide-time'}."(".$val->{'tide-level'}."cm)";    // for debug

                        // 始めの位置が満潮か干潮かによってexcelでの表示位置をstring化
                        $highString = $high["$highCount"].$day_count;
//                        echo $highString;         // for debug
                        $sheet->setCellValue("$highString",$val->{'tide-time'}."(".$val->{'tide-level'}.")");
                    }
                    $highCount++;
                }
                $sheet->getStyle( "A$day_count:G$day_count" )->getBorders()->getBottom()->setBorderStyle( PHPExcel_Style_Border::BORDER_THIN );
                $day_count++;
            }
            $i++;
            $sheet->getStyle( 'A1:G40' )->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle( 'A1:G40' )->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $sheet->getStyle( 'A1:G4' )->getFont()->setSize( 26 );
            $day_count--;
            $sheet->getStyle( "A5:G$day_count" )->getBorders()->getTop()->setBorderStyle( PHPExcel_Style_Border::BORDER_THIN );
            $sheet->getStyle( "A5:G5" )->getBorders()->getBottom()->setBorderStyle( PHPExcel_Style_Border::BORDER_THIN );
            $sheet->getStyle( "A6:G6" )->getBorders()->getBottom()->setBorderStyle( PHPExcel_Style_Border::BORDER_THIN );
            $sheet->getStyle( "A5:G$day_count" )->getBorders()->getBottom()->setBorderStyle( PHPExcel_Style_Border::BORDER_THIN );
            $sheet->getStyle( "A5:A$day_count" )->getBorders()->getLeft()->setBorderStyle( PHPExcel_Style_Border::BORDER_THIN );
            $sheet->getStyle( "G5:G$day_count" )->getBorders()->getRight()->setBorderStyle( PHPExcel_Style_Border::BORDER_THIN );
            $sheet->getStyle( "A6:A$day_count" )->getBorders()->getRight()->setBorderStyle( PHPExcel_Style_Border::BORDER_DOTTED );
            $sheet->getStyle( "B6:B$day_count" )->getBorders()->getRight()->setBorderStyle( PHPExcel_Style_Border::BORDER_DOTTED );
            $sheet->getStyle( "C6:C$day_count" )->getBorders()->getRight()->setBorderStyle( PHPExcel_Style_Border::BORDER_DOTTED );
            $sheet->getStyle( "D7:D$day_count" )->getBorders()->getRight()->setBorderStyle( PHPExcel_Style_Border::BORDER_DOTTED );
            $sheet->getStyle( "E6:E$day_count" )->getBorders()->getRight()->setBorderStyle( PHPExcel_Style_Border::BORDER_DOTTED );
            $sheet->getStyle( "F7:F$day_count" )->getBorders()->getRight()->setBorderStyle( PHPExcel_Style_Border::BORDER_DOTTED );
            $sheet->getStyle( "A1:G$day_count")->getFont()->setBold(true);
        }
    }

    $fname = $portname[$form2date->getPortid()];
    $filename = "tide_" . $fname . ".xlsx";
    // Excel2007形式で出力する
    header('Content-Type: application/vnd.ms-excel');
    header("Content-Disposition: attachment;filename='$filename'");
    header('Cache-Control: max-age=0');

    $writer = PHPExcel_IOFactory::createWriter($book, "Excel2007");
    $writer->save('php://output');


    echo "0";




} else {
    echo "このページへは直接アクセスできません。";
    exit(1);

}


?>

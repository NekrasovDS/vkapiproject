<doctype html>
<html>
<title></title>
<head>Люди и группы:<br></head>
<body>
<?php
    require 'vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Style\Color;
    use PhpOffice\PhpSpreadsheet\Style\ConditionalFormatting\Wizard;
    use PhpOffice\PhpSpreadsheet\Style\Fill;
    use PhpOffice\PhpSpreadsheet\Style\Style;
    set_time_limit(0);

    $spreadsheet = new Spreadsheet();
    $activeWorksheet = $spreadsheet->getActiveSheet();
    //$activeWorksheet->setCellValue('A1', 'Hello World !');
    $activeWorksheet->setCellValue('A1', 'Имя')
        ->setCellValue('B1', 'Фамилия')
        ->setCellValue('C1', 'Id пользователя')
        ->setCellValue('D1', 'Id группы')
        ->setCellValue('E1', 'Направление');
    
    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(15);
    $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(15);
    $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(18);
    $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
    $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(30);
    
    $headerStyles = [
    'font'=>[
      'color'=>[
        'rgb' => '000'
      ],
      'bold' => true,
      'size' => 13
    ],
    'fill'=>[
        'fillType' => Fill::FILL_SOLID,
        'startColor' => [
          'rgb' => Color::COLOR_RED
            ]
        ],
    ];
    
    $spreadsheet->getActiveSheet()->getStyle('A1:E1')->applyFromArray($headerStyles);
    
    $user_id = 0; //Вставьте свой id
    $count = 1;
    $token = ""; //Вставьте свой токен
    
    $request_params = array(
        'user_id' => $user_id,
        'order' => 'random',
        'count' => $count,
        'v' => '5.199',
        'fields' => 'first_name,last_name',
        'access_token' => $token
    );
    $num = 1;
    
    $get_params = http_build_query($request_params);
    for ($i = 1; $i <= 75; $i++){
                  
        $friends = json_decode(file_get_contents('http://api.vk.com/method/friends.get?' .$get_params), true);
        foreach($friends['response']['items'] as $item1){
            echo $item1["first_name"], " ", $item1["last_name"], " (";
            echo $item1["id"], "): <br>";
            $user_id = $item1["id"];
            
            $num++;
            $activeWorksheet->setCellValue('A'. $num, $item1["first_name"])
                ->setCellValue('B'. $num, $item1["last_name"])
                ->setCellValue('C'. $num, $item1["id"]);
            $groups = json_decode(file_get_contents("https://api.vk.com/method/users.getSubscriptions?user_id=$user_id&extended=1&count=25&v=5.199&access_token=" .$token), true);
            foreach($groups['response']['items'] as $item2){
                $group_id = $item2["id"];
                $group_activity = json_decode(file_get_contents("http://api.vk.com/method/groups.getById?group_id=$group_id&fields=activity&v=5.199&access_token=" .$token), true);
                foreach($group_activity['response']['groups'] as $item3){
                    echo $item3['activity'], " ";
                    $activeWorksheet->setCellValue('E'. $num, $item3['activity']);
                }
                echo $group_id, "<br>";
                $activeWorksheet->setCellValue('D'. $num, $item2["id"]);
                $num++;
            }
    }
    echo "<br>";
    }
    
    $datetime = date(d.m.Y);
    $writer = new Xlsx($spreadsheet);
    $writer->save("user_database_$datetime.xlsx");
    //print_r($friends);
    //print_r($query);
?>
</body>
</html>
</doctype>

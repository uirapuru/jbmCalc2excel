<?php
include "vendor/autoload.php";
use Goutte\Client;
use PhpOffice\PhpSpreadsheet\IOFactory;

$temps = ["first" => 5, "second" => 15, "third" => 25];
$windSpeeds = ["first" => 2, "second" => 5, "third" => 10];
$movTargetSpeeds = ["first" => 1, "second" => 2, "third" => 5];

$input = [
    "b_id.v" => "1084",
    "bc.v" => "0.505",
    "d_f.v" => "0",
    "bt_wgt.v" => "175",
    "bt_wgt.u" => "23",
    "cal.v" => "0.308",
    "cal.u" => "8",
    "blt_len.v" => "1.24",
    "blt_len.u" => "8",
    "tip_len.v" => "0.0",
    "tip_len.u" => "8",
    "m_vel.v" => "820",
    "m_vel.u" => "17",
    "ch_dst.v" => "10",
    "ch_dst.u" => "11",
    "hgt_sgt.v" => "5",
    "hgt_sgt.u" => "11",
    "ofs_sgt.v" => "0.0",
    "ofs_sgt.u" => "8",
    "hgt_zer.v" => "0.0",
    "hgt_zer.u" => "8",
    "ofs_zer.v" => "0.0",
    "ofs_zer.u" => "8",
    "azm.v" => "0.0",
    "azm.u" => "4",
    "ele.v" => "0.0",
    "ele.u" => "4",
    "los.v" => "0.0",
    "cnt.v" => "0.0",
    "b_twt.v" => "12.0",
    "b_twt.u" => "8",
    "b_twt_dir.v" => "1",
    "spd_wnd.v" => "0",
    "spd_wnd.u" => "17",
    "ang_wnd.v" => "270",
    "spd_tgt.v" => "5",
    "spd_tgt.u" => "17",
    "ang_tgt.v" => "90.0",
    "siz_tgt.v" => "50",
    "siz_tgt.u" => "11",
    "rng_min.v" => "100",
    "rng_max.v" => "1200",
    "rng_inc.v" => "25",
    "rng_zer.v" => "100",
    "tmp.v" => "15",
    "tmp.u" => "18",
    "prs.v" => "1000",
    "prs.u" => "21",
    "hum.v" => "50",
    "alt.v" => "100",
    "alt.u" => "12",
    "cor_prs.v" => "on",
    "rad_vz.v" => "20",
    "rad_vz.u" => "11",
    "col_eng.v" => "1",
    "col1_un.v" => "1.00",
    "col1_un.u" => "4",
    "col2_un.v" => "1.00",
    "col2_un.u" => "11",
    "cor_ele.v" => "on",
    "rng_un.v" => "on",
    "def_cnt.v" => "on",
    "mrk_trs.v" => "on",
    "ext_row.v" => "on",
    "inc_drf.v" => "on",
    "inc_ds.v" => "on",
];

function getOutput(array $input) {
    $client = new Client();
    $crawler = $client->request('GET', 'https://www.jbmballistics.com/cgi-bin/jbmtraj_drift-5.1.cgi');

    $form = $crawler->filter("form")->form();

    $resp = $client->submit($form, $input);
    $table = $resp->filter("table.output_table")->first();

    $extr = function (string $col, int $eq) {
        return function ($tr, $i) use ($col, $eq) {
            return $tr->filter($col)->eq($eq)->each(fn($td, $i) => trim($td->text()));
        };
    };

    return [
        "range" => array_filter(array_merge(...$table->filter("tr")->each($extr('td.range_cell', 0)))),
        "elevation" => array_filter(array_merge(...$table->filter("tr")->each($extr('td.drop_angle_cell', 0)))),
        "windage" => array_filter(array_merge(...$table->filter("tr")->each($extr('td.wind_angle_cell', 0)))),
        "lead" => array_filter(array_merge(...$table->filter("tr")->each($extr('td.lead_angle_cell', 0)))),
    ];
}

$data = [
    "temps" => array_map(fn($temp) => getOutput(array_merge($input, ["tmp.v" => $temp])), $temps),
    "wind_r" => array_map(fn($temp) => array_map(fn($wind) => getOutput(array_merge($input, ["tmp.v" => $temp, "spd_wnd.v" => $wind, "ang_wnd.v" => 90])), $windSpeeds), $temps),
    "wind_l" => array_map(fn($temp) => array_map(fn($wind) => getOutput(array_merge($input, ["tmp.v" => $temp, "spd_wnd.v" => $wind, "ang_wnd.v" => 270])), $windSpeeds), $temps),
    "move" => array_map(fn($temp) => array_map(fn($speed) => getOutput(array_merge($input, ["tmp.v" => $temp, "spd_tgt.v" => $speed])), $movTargetSpeeds), $temps),
];

$spreadsheet = IOFactory::load("template.xlsx");
$sheet = $spreadsheet->getActiveSheet();

$sheet->fromArray(array_chunk($data["temps"]["first"]["elevation"], 1), NULL, "B4");
$sheet->fromArray(array_chunk($data["temps"]["second"]["elevation"], 1), NULL, "C4");
$sheet->fromArray(array_chunk($data["temps"]["third"]["elevation"], 1), NULL, "D4");
$sheet->fromArray(array_chunk($data["temps"]["first"]["windage"], 1), NULL, "E4");
$sheet->fromArray(array_chunk($data["temps"]["second"]["windage"], 1), NULL, "F4");
$sheet->fromArray(array_chunk($data["temps"]["third"]["windage"], 1), NULL, "G4");

$sheet->fromArray(array_chunk($data["wind_r"]["first"]["first"]["windage"], 1), NULL, "H4");
$sheet->fromArray(array_chunk($data["wind_r"]["first"]["second"]["windage"], 1), NULL, "I4");
$sheet->fromArray(array_chunk($data["wind_r"]["first"]["third"]["windage"], 1), NULL, "J4");
$sheet->fromArray(array_chunk($data["wind_r"]["second"]["first"]["windage"], 1), NULL, "K4");
$sheet->fromArray(array_chunk($data["wind_r"]["second"]["second"]["windage"], 1), NULL, "L4");
$sheet->fromArray(array_chunk($data["wind_r"]["second"]["third"]["windage"], 1), NULL, "M4");
$sheet->fromArray(array_chunk($data["wind_r"]["third"]["first"]["windage"], 1), NULL, "N4");
$sheet->fromArray(array_chunk($data["wind_r"]["third"]["second"]["windage"], 1), NULL, "O4");
$sheet->fromArray(array_chunk($data["wind_r"]["third"]["third"]["windage"], 1), NULL, "P4");

$sheet->fromArray(array_chunk($data["wind_l"]["first"]["first"]["windage"], 1), NULL, "Q4");
$sheet->fromArray(array_chunk($data["wind_l"]["first"]["second"]["windage"], 1), NULL, "R4");
$sheet->fromArray(array_chunk($data["wind_l"]["first"]["third"]["windage"], 1), NULL, "S4");
$sheet->fromArray(array_chunk($data["wind_l"]["second"]["first"]["windage"], 1), NULL, "T4");
$sheet->fromArray(array_chunk($data["wind_l"]["second"]["second"]["windage"], 1), NULL, "U4");
$sheet->fromArray(array_chunk($data["wind_l"]["second"]["third"]["windage"], 1), NULL, "V4");
$sheet->fromArray(array_chunk($data["wind_l"]["third"]["first"]["windage"], 1), NULL, "W4");
$sheet->fromArray(array_chunk($data["wind_l"]["third"]["second"]["windage"], 1), NULL, "X4");
$sheet->fromArray(array_chunk($data["wind_l"]["third"]["third"]["windage"], 1), NULL, "Y4");

$sheet->fromArray(array_chunk($data["move"]["first"]["first"]["lead"], 1), NULL, "Z4");
$sheet->fromArray(array_chunk($data["move"]["first"]["second"]["lead"], 1), NULL, "AA4");
$sheet->fromArray(array_chunk($data["move"]["first"]["third"]["lead"], 1), NULL, "AB4");
$sheet->fromArray(array_chunk($data["move"]["second"]["first"]["lead"], 1), NULL, "AC4");
$sheet->fromArray(array_chunk($data["move"]["second"]["second"]["lead"], 1), NULL, "AD4");
$sheet->fromArray(array_chunk($data["move"]["second"]["third"]["lead"], 1), NULL, "AE4");
$sheet->fromArray(array_chunk($data["move"]["third"]["first"]["lead"], 1), NULL, "AF4");
$sheet->fromArray(array_chunk($data["move"]["third"]["second"]["lead"], 1), NULL, "AG4");
$sheet->fromArray(array_chunk($data["move"]["third"]["third"]["lead"], 1), NULL, "AH4");

$writer = IOFactory::createWriter($spreadsheet, "Xlsx");
$writer->save("result_".time().".xlsx");

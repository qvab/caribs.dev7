<?php

namespace frontend\controllers;

use Composer\Package\Package;
use Faker\Provider\Base;
use function foo\func;
use frontend\models\Excursions;
use frontend\models\Language;
use frontend\models\PayArea;
use frontend\models\Partners;
use frontend\models\MethodComunications;
use frontend\models\Settings;
use frontend\models\Amocrm;
use frontend\models\AmoSettings;
use frontend\models\PayMethod;
use frontend\models\HotelExcursions;
use function GuzzleHttp\Psr7\str;
use Yii;
use frontend\models\Records;
use frontend\models\Hotels;
use frontend\models\RecordsSearch;
use yii\helpers\ArrayHelper;
use yii\rest\ActiveController;
use yii\web\Controller;
use yii\web\NotFoundHttpException;
use yii\filters\VerbFilter;
use frontend\models\ExcursionsDate;

include_once $_SERVER["DOCUMENT_ROOT"]."/yii/frontend/views/report-amo/amo.class.php";
include_once $_SERVER["DOCUMENT_ROOT"]."/linerapp/miwr/classes/class.records.php";
set_time_limit(3600);

/**
 *
 * ExcursionsController implements the CRUD actions for Excursions model.
 */
class ReportAmoController extends Controller
{
  public $modelClass = 'frontend\models\App';
  private $obAmo;
  private $arDate = [];
  private
    $arAllLeads = [],
    $arContacts = [],
    $arContactsDouble = [],
    $arContactIds = [],
    $arSuccessLeads = [
    "all" => 0,
    "dates" => []
  ],
    $arTargetCount = [
    "all" => 0,
    "dates" => []
  ],
    $arConversions = [
    "all" => 0,
    "middle" => 0,
    "dates" => []
  ];

  private $iSizeAllLeads = 0;

  private $arTargetsForArea = [
    "republica" => [
      "all" => 0,
      "dates" => []
    ],
    "ilya" => [
      "all" => 0,
      "dates" => []
    ],
  ];

  private $arCancel = [
    "Другой регион",
    "СПАМ",
    "Нет контактов",
    "Туристы не прилетели",
    "Купили в пляжном офисе",
    "Вопрос не по экскурсиям"
  ];

  private $arTagsCancel = [
    "Дубликат",
    "турагентство"
  ];

  private $arContactIdsPages = [];

  public function actionIndex()
  {

    return $this->render('index', [
      'searchModel' => "test",
    ]);
  }

  /**
   * Группирирующий формат даты
   */
  private function __d($d)
  {
    return date("d.m.Y", $d);
  }

  /**
   * Расчет колонки конверсии
   */
  private function renumberConversion()
  {
    $iAllPercent = 0;
    foreach ($this->arDate["list"] as $date) {
      if (!empty($this->arSuccessLeads["dates"][$date])) {
        $this->arConversions["dates"][$date] = (float)number_format($this->arSuccessLeads["dates"][$date] / ($this->arTargetCount["dates"][$date] / 100), 2);
      } else {
        $this->arConversions["dates"][$date] = 0;
      }
      $iAllPercent += $this->arConversions["dates"][$date];
    }
    $this->arConversions["all"] = (float)number_format($this->arSuccessLeads["all"] / ($this->arTargetCount["all"] / 100), 2);
    $this->arConversions["middle"] = (float)number_format($iAllPercent / $this->arDate["count_list"], 2);

  }

  /**
   * Иницилизируем данные
   */
  private function __init()
  {

    $this->obAmo = new \MW_AMO_CRM();
    $this->obAmo->init("republicapro", "consultant@linerapp.com", "94073ea1d7d3092a4facb00bf6ea5ff6ae49fffc");


    $this->arDate["list"] = [];
    $this->arDate["from"] = strtotime($_REQUEST["date_from"]);
    $this->arDate["to"] = strtotime($_REQUEST["date_to"]." + 1 day") - 1;

    $this->arDate["count_day"] = cal_days_in_month(CAL_GREGORIAN,
      date("m", $this->arDate["from"]),
      date("Y", $this->arDate["from"]));
    for ($d = (int)date("d", $this->arDate["from"]); $d <= (int)date("d", $this->arDate["to"]); $d++) {
      $sDay = $d < 10 ? "0".$d : $d;
      $sKey = date("d.m.Y", strtotime($sDay.".".date("m.Y", $this->arDate["from"])));
      $this->arDate["list"][] = $sKey;
      $this->arTargetCount["dates"][$sKey] = 0;
    }

    $this->arDate["count_list"] = count($this->arDate["list"]);

  }


  /**
   * Выбор сделок через AJAX
   * @return string
   */
  public function actionSelectedLeads()
  {
    $this->__init();
    /**
     * Выборка сделок
     */
    $iNext = 0;
    $arLeadsFilter = [];
    $iStart = $_POST["offset"];

    $arLeads = $this->obAmo->getLeadList([
      "filter" =>
        [
          "date_create" => [
            "from" => $this->arDate["from"],
            "to" => $this->arDate["to"]
          ]
        ],
      "limit_rows" => 500,
      "limit_offset" => $iStart
    ]);

    if (!empty($arLeads["response"]["_embedded"]["items"])) {
      $arLeadsAll = $arLeads["response"]["_embedded"]["items"];
      $iCount = count($arLeadsAll);
      if ($iCount > 499) {
        $iNext = 1;
      }

      foreach ($arLeadsAll as $arLead) {
        $sKeyDate = $this->__d($arLead["created_at"]); // Ключ даты
        if (empty($sKeyDate)) {
          vd2($arLead);
        }
        /**
         * Основые данные сделки
         */
        $arLeadsFilter[$arLead["id"]] = [
          "id" => $arLead["id"],
          "name" => $arLead["name"],
          "roistat" => false,
          "cancel" => false,
          "tags" => [],
          "contacts" => $arLead["contacts"]["id"],
          "status_id" => $arLead["status_id"],
          "date_period" => $sKeyDate,
          "created_at" => $arLead["created_at"]
        ];


        /**
         * Выбираем только названия тегов
         * И отменяем выбор такой сделки
         */
        if (!empty($arLead["tags"])) {
          foreach ($arLead["tags"] as $arTag) {
            if (in_array($arTag["name"], $this->arTagsCancel)) {
              unset($arLeadsFilter[$arLead["id"]]);
              continue 2;
            }
            $arLeadsFilter[$arLead["id"]]["tags"][] = $arTag["name"];
          }
        }

        /**
         * Выбираем кастомные поля, которые нам нужны
         */
        foreach ($arLead["custom_fields"] as $arField) {
          switch ($arField["name"]) {
            case 'roistat':
              $arLeadsFilter[$arLead["id"]]["roistat"] = $arField["values"][0]["value"];
              break;
            case 'Причина отказа':
              if (in_array($arField["values"][0]["value"], $this->arCancel)) {
                unset($arLeadsFilter[$arLead["id"]]);
                continue 3;
              }
              $arLeadsFilter[$arLead["id"]]["cancel"] = $arField["values"][0]["value"];
              break;
          }
        }

      }
      return json_encode([
        "error" => 0,
        "leads" => $arLeadsFilter,
        "next" => $iNext,
        "count" => $iCount
      ], JSON_UNESCAPED_UNICODE);
    }

    return json_encode(["error" => 1]);
  }


  /**
   * Выборка контактов AJAX
   */
  public function actionSelectedContacts()
  {

    $this->__init();
    /**
     * Выборка контактов
     */
    $arAllContactsSelect = [];
    $arDoubleEmails = [];
    $arDoublePhones = [];
    $arEmails = [];
    $arPhones = [];
    $iNext = 0;
    $arLeadsFilter = [];
    $arLeadsAll = [];
    $this->arAllLeads = json_decode($_REQUEST["leads"], true);
    $iLeadStart = count($this->arAllLeads);
    $arLeads = $this->obAmo->getContactsList([
      "id" => json_decode($_REQUEST["ids"], true)
    ]);

    if (!empty($arLeads["response"]["_embedded"]["items"])) {
      $arLeadsAll = $arLeads["response"]["_embedded"]["items"];
      $iCount = count($arLeadsAll);
      if ($iCount > 498) {
        $iNext = 1;
      }


      foreach ($arLeadsAll as $arContact) {
        foreach ($arContact["custom_fields"] as $arField) {
          switch ($arField["code"]) {
            case "PHONE":
              $arPhones[$arField["values"][0]["value"]]["id"][] = $arContact["id"];
              $arPhones[$arField["values"][0]["value"]]["leads"]["contacts"][$arContact["id"]] = $arContact["leads"]["id"];
              if (count($arPhones[$arField["values"][0]["value"]]["id"]) > 1) {
                $arDoublePhones[$arField["values"][0]["value"]] = $arPhones[$arField["values"][0]["value"]];
              }
              break;
            case "EMAIL":
              $arEmails[$arField["values"][0]["value"]]["id"][] = $arContact["id"];
              $arEmails[$arField["values"][0]["value"]]["leads"]["contacts"][$arContact["id"]] = $arContact["leads"]["id"];
              if (count($arEmails[$arField["values"][0]["value"]]["id"]) > 1) {
                $arDoubleEmails[$arField["values"][0]["value"]] = $arEmails[$arField["values"][0]["value"]];
              }
              break;
          }
        }
      }


      foreach ($arDoublePhones as $querySearch => $arDouble) {
        foreach ($arDouble["id"] as $id) {
          foreach ($arDoublePhones[$querySearch]["leads"]["contacts"][$id] as $idLead) {
            if (!empty($this->arAllLeads[$idLead])) {
              $arAllContactsSelect[] = $idLead;
              unset($this->arAllLeads[$idLead]);
            }
          }
        }
      }
      return json_encode([
        "error" => 0,
        //"leads_id" => $arAllContactsSelect,
        "next" => $iNext,
        "count" => $iCount,
        "all_leads" => $this->arAllLeads,
        "all_leads_start" => $iLeadStart,
        "all_leads_count" => count($this->arAllLeads),
      ], JSON_UNESCAPED_UNICODE);
    }

    return json_encode(["error" => 1]);

  }

  /**
   * Сортировка контактов и посик дублей
   * @return string
   */
  public function actionFilterContacts()
  {
    $arLeadsFilter = json_decode($_REQUEST["leads"], true);
    $iStart = count($arLeadsFilter);
    foreach ($arLeadsFilter as $arLead) {
      if (!empty($arLead["contacts"])) {
        foreach ($arLead["contacts"] as $idContact) {
          $this->arContacts[$idContact][] = $arLead["id"];
          if (count($this->arContacts[$idContact]) > 1) {
            $this->arContactsDouble[$idContact] = $this->arContacts[$idContact];
            unset($arLeadsFilter[$arLead["id"]]);
            unset($this->arContacts[$idContact]);
            continue 2;
          }
          if (!in_array($idContact, $this->arContactIds)) {
            $this->arContactIds[] = $idContact;
            $iPage = floor(count($this->arContactIds) / 500);
            $this->arContactIdsPages[$iPage][] = $idContact;
          }
        }
      }
    }

    return json_encode([
      "error" => 0,
      "contacts_ids" => $this->arContactIds,
      "contacts_count_ids" => count($this->arContactIds),
      "pages" => $this->arContactIdsPages,
      "count" => count($arLeadsFilter),
      "start" => $iStart
    ], JSON_UNESCAPED_UNICODE);

  }


  /**
   * Считаем целевые обращения AJAX
   */
  public function actionRenumberTarget()
  {
    $this->__init();
    $this->arAllLeads = json_decode($_REQUEST["leads"], true);
    foreach ($this->arAllLeads as $idLead => $arLead) {
      $sKeyDate = $arLead["date_period"]; // Ключ даты
      /**
       * Записываем по месту продажи
       */
      if (!empty($arLead["roistat"])) {
        $sKeyArea = false;
        if ($arLead["roistat"] == "ilya") {
          $sKeyArea = "ilya";
        } else {
          $sKeyArea = "republica";
        }

        // Производим запись
        if (!empty($sKeyArea)) {
          $this->arTargetsForArea[$sKeyArea]["all"]++;
          $this->arTargetsForArea[$sKeyArea]["middle"] = (float)number_format($this->arTargetsForArea[$sKeyArea]["all"] / $this->arDate["count_list"], 2);

          if (!empty($this->arTargetsForArea[$sKeyArea]["dates"][$sKeyDate])) {
            $this->arTargetsForArea[$sKeyArea]["dates"][$sKeyDate]++;
          } else {
            $this->arTargetsForArea[$sKeyArea]["dates"][$sKeyDate] = 1;
          }
        }
      }

      /**
       * Проверка на успешно реализованную сделку
       */

      if ($arLead["status_id"] == 142) {
        $arLead["status_success"] = true;
        $this->arSuccessLeads["all"]++;
        $this->arSuccessLeads["middle"] = (float)number_format($this->arSuccessLeads["all"] / $this->arDate["count_list"], 2);
        if (!empty($this->arSuccessLeads["dates"][$sKeyDate])) {
          $this->arSuccessLeads["dates"][$sKeyDate]++;
        } else {
          $this->arSuccessLeads["dates"][$sKeyDate] = 1;
        }

      } else {
        $this->arAllLeads[$idLead]["status_success"] = false;
      }
      // Записываем количество целевых обращений ВСЕГО по датам
      $this->arTargetCount["dates"][$sKeyDate]++;
      $this->arTargetCount["all"]++;
      $this->arTargetCount["middle"] = (float)number_format($this->arTargetCount["all"] / $this->arDate["count_list"], 2);
    }

    $this->renumberConversion();

    $arMoneys = $this->getTickets();
    return json_encode([
      "period" => $this->arDate["list"],
      "count_leads" => count($this->arAllLeads),
      "arTargetCount" => $this->arTargetCount,
      "arSuccessLeads" => $this->arSuccessLeads,
      "arTargetsForArea" => $this->arTargetsForArea,
      "arConversions" => $this->arConversions,
      "moneys" => $arMoneys
    ], JSON_UNESCAPED_UNICODE);
  }


  /**
   * Выборка выписанных билетов
   * Не резер и не отмененный
   * @return array
   */
  private function getTickets()
  {
    $obRecord = new \MIWR\Records();
    $sSQL = "SELECT 
records.c_total,
records.price_discount,
records.pay_area_id,
records.price AS price_real,
IF(records.price_discount IS NOT NULL, round(records.price * ( (100 - records.price_discount) / 100)), records.price) AS price,
tickets.date_created AS created
FROM records 
LEFT JOIN tickets ON records.ticket = tickets.number 
WHERE (tickets.date_created >= ".$this->arDate["from"]." AND tickets.date_created <= ".$this->arDate["to"].") AND is_cancel = 0
ORDER BY tickets.date_created ASC";

    $arMonthsForAreaData = [];
    $arMonthsForTotalData = [];
    $iTotalPrice = 0;
    $obRecord->dbQuery($sSQL);
    while ($arTicket = $obRecord->fetch()) {
      $keyDate = $this->__d($arTicket["created"]);
      $keyArea = false;
      switch ($arTicket["pay_area_id"]) {
        case 7:
          $keyArea = "ilya";
          break;
        case 5:
          $keyArea = "playa";
          break;
        case 6:
          $keyArea = "cabeza";
          break;
        case 4:
          $keyArea = "republica";
          break;
      }
      if (!empty($keyArea)) {
        $iTotalPrice += $arTicket["price"];
        if (!empty($arMonthsForAreaData["dates"][$keyDate][$keyArea]["price"])) {
          $arMonthsForAreaData["dates"][$keyDate][$keyArea]["price"] += $arTicket["price"];
        } else {
          $arMonthsForAreaData["dates"][$keyDate][$keyArea]["price"] = $arTicket["price"];
        }


        if (!empty($arMonthsForAreaData["dates"][$keyDate][$keyArea]["people"])) {
          $arMonthsForAreaData["dates"][$keyDate][$keyArea]["people"] += $arTicket["c_total"];
        } else {
          $arMonthsForAreaData["dates"][$keyDate][$keyArea]["people"] = $arTicket["c_total"];
        }

        // общие значения

        if (!empty($arMonthsForAreaData["all"][$keyArea]["price"])) {
          $arMonthsForAreaData["all"][$keyArea]["price"] += $arTicket["price"];
          $arMonthsForAreaData["middle"][$keyArea]["price"] = ($arMonthsForAreaData["all"][$keyArea]["price"] / $this->arDate["count_list"]);
        } else {
          $arMonthsForAreaData["all"][$keyArea]["price"] = $arTicket["price"];
        }

        if (!empty($arMonthsForAreaData["all"][$keyArea]["people"])) {
          $arMonthsForAreaData["all"][$keyArea]["people"] += $arTicket["c_total"];
          $arMonthsForAreaData["middle"][$keyArea]["people"] = ($arMonthsForAreaData["all"][$keyArea]["people"] / $this->arDate["count_list"]);
        } else {
          $arMonthsForAreaData["all"][$keyArea]["people"] = $arTicket["c_total"];
        }


        /**
         * Считаем итоговые значения
         */
        if (!empty($arMonthsForTotalData["dates"][$keyDate]["all_price"])) {
          $arMonthsForTotalData["dates"][$keyDate]["all_price"] += $arTicket["price"];
        } else {
          $arMonthsForTotalData["dates"][$keyDate]["all_price"] = $arTicket["price"];
        }

        if (!empty($arMonthsForTotalData["dates"][$keyDate]["all_people"])) {
          $arMonthsForTotalData["dates"][$keyDate]["all_people"] += $arTicket["c_total"];
        } else {
          $arMonthsForTotalData["dates"][$keyDate]["all_people"] = $arTicket["c_total"];
        }

        if (!empty($arMonthsForTotalData["dates"][$keyDate]["fact_price"])) {
          $arMonthsForTotalData["dates"][$keyDate]["fact_price"] += $arTicket["price"];
        } else {
          $arMonthsForTotalData["dates"][$keyDate]["fact_price"] = $iTotalPrice;
        }

      }
    }

    // Вычисляем прогноз продаж и средний чек
    foreach ($arMonthsForTotalData["dates"] as $keyDate => $arData) {
      $arMonthsForTotalData["dates"][$keyDate]["forecast"] = $arData["all_price"] * $this->arDate["count_day"];
      $arMonthsForTotalData["dates"][$keyDate]["middle_check"] = (float)number_format($arData["all_price"] / $arData["all_people"], 2);

      if (!empty($arMonthsForTotalData["all"]["forecast"])) {
        $arMonthsForTotalData["all"]["forecast"] += $arMonthsForTotalData["dates"][$keyDate]["forecast"];
        $arMonthsForTotalData["all"]["middle_check"] += $arMonthsForTotalData["dates"][$keyDate]["middle_check"];
      } else {
        $arMonthsForTotalData["all"]["forecast"] = $arMonthsForTotalData["dates"][$keyDate]["forecast"];
        $arMonthsForTotalData["all"]["middle_check"] = $arMonthsForTotalData["dates"][$keyDate]["middle_check"];
      }

      if (!empty($arMonthsForTotalData["all"]["all_price"])) {
        $arMonthsForTotalData["all"]["all_price"] += $arData["all_price"];
        $arMonthsForTotalData["all"]["all_people"] += $arData["all_people"];
        $arMonthsForTotalData["all"]["fact_price"] += $arData["fact_price"];
      } else {
        $arMonthsForTotalData["all"]["all_price"] = $arData["all_price"];
        $arMonthsForTotalData["all"]["all_people"] = $arData["all_people"];
        $arMonthsForTotalData["all"]["fact_price"] = $arData["fact_price"];
      }
    }


    $arMonthsForTotalData["middle"]["all_price"] = $arMonthsForTotalData["all"]["all_price"] / $this->arDate["count_list"];
    $arMonthsForTotalData["middle"]["all_people"] = $arMonthsForTotalData["all"]["all_people"] / $this->arDate["count_list"];
    $arMonthsForTotalData["middle"]["fact_price"] = $arMonthsForTotalData["all"]["fact_price"] / $this->arDate["count_list"];
    $arMonthsForTotalData["middle"]["middle_check"] = $arMonthsForTotalData["all"]["middle_check"] / $this->arDate["count_list"];
    $arMonthsForTotalData["middle"]["forecast"] = $arMonthsForTotalData["all"]["forecast"] / $this->arDate["count_list"];

    return [
      "arMonthsDate" => $arMonthsForAreaData,
      "arMonthsForTotalData" => $arMonthsForTotalData
    ];
  }

  public function beforeAction($action)
  {
    if ($action->id == 'generate-today') {
      Yii::$app->controller->enableCsrfValidation = false;
    }

    return parent::beforeAction($action);
  }

  /**
   * Генерация ежедневного отчета
   */
  public function actionGenerateToday()
  {
    $this->__init();
    $this->enableCsrfValidation = false;
    header("Content-type: text/html; charset=UTF-8");
    $obExcel = new \PHPExcel();
    $obExcel->setActiveSheetIndex(0);
    $sheet = $obExcel->getActiveSheet();
    $sheet->setTitle('Ежедневный отчет');

    $arTitlesOne = [
      "A" => "",
      "B" => "Playa",
      "E" => "Кабеса",
      "H" => "Республика",
      "K" => "Илья",
      "N" => "ИТОГО",
      "S" => "ПЛАН ПО ВЫРУЧКЕ",
      "T" => $_GET["plan"],
    ];
    $arTitlesTwo = [
      "A" => "Дата",
      "B" => "Целевые обр",
      "C" => "Людей",
      "D" => "Выручка",
      "E" => "Целевые обр",
      "F" => "Людей",
      "G" => "Выручка",
      "H" => "Целевые обр",
      "I" => "Людей",
      "J" => "Выручка",
      "K" => "Целевые обр",
      "L" => "Людей",
      "M" => "Выручка",
      "N" => "Целевые обр",
      "O" => "Люди",
      "P" => "Выручка",
      "Q" => "Средний чек",
      "R" => "% конверсии\n(отношение положительно\nзакрытых ко всем закрытым целевым)",
      "S" => "факт общая\nвыручка на\nтекущую дату",
      "T" => "прогноз на\nконец месяца",
    ];

    $sheet->mergeCells("B1:D1");
    $sheet->mergeCells("E1:G1");
    $sheet->mergeCells("H1:J1");
    $sheet->mergeCells("K1:M1");
    $sheet->mergeCells("N1:R1");

    foreach ($arTitlesOne as $key => $val) {
      $sheet->setCellValue($key."1", $val);
      $sheet->getStyle($key."1")->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
      $sheet->getStyle($key."1")->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
      $sheet->getStyle($key."1")->getAlignment()->setWrapText(true);
      $sheet->getStyle($key."1")->applyFromArray([
        'fill' => [
          'type' => \PHPExcel_Style_Fill::FILL_SOLID,
          'color' => ['rgb' => 'FFFF00'],
        ]
      ]);
      $sheet->getColumnDimension($key)->setAutoSize(true);
    }
    foreach ($arTitlesTwo as $key => $val) {
      $sheet->setCellValue($key."2", $val);
      $sheet->getStyle($key."2")->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
      $sheet->getStyle($key."2")->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
      $sheet->getStyle($key."2")->getAlignment()->setWrapText(true);
      $sheet->getStyle($key."2")->applyFromArray([

        'fill' => [
          'type' => \PHPExcel_Style_Fill::FILL_SOLID,
          'color' => ['rgb' => 'FFFF00']
        ],
        "borders" => [
          'bottom' => [
            'style' => \PHPExcel_Style_Border::BORDER_MEDIUM,
            'color' => ['rgb' => '808080']
          ],
          'top' => [
            'style' => \PHPExcel_Style_Border::BORDER_MEDIUM,
            'color' => ['rgb' => '808080']
          ],
          'left' => [
            'style' => \PHPExcel_Style_Border::BORDER_MEDIUM,
            'color' => ['rgb' => '808080']
          ],
          'right' => [
            'style' => \PHPExcel_Style_Border::BORDER_MEDIUM,
            'color' => ['rgb' => '808080']
          ]
        ]
      ]);
      $sheet->getColumnDimension($key)->setAutoSize(true);
    }
    $sheet->setCellValue("A3", "Среднее показатель");


    $arData = json_decode($_POST["data"], true);


    /**
     * Пишем средние показатели
     */

    $sheet->setCellValue("H3", $arData["arTargetsForArea"]["republica"]["middle"]);
    $sheet->setCellValue("K3", $arData["arTargetsForArea"]["ilya"]["middle"]);

    $sheet->setCellValue("C3", $arData["moneys"]["arMonthsDate"]["middle"]["playa"]["people"]);
    $sheet->setCellValue("F3", $arData["moneys"]["arMonthsDate"]["middle"]["cabeza"]["people"]);
    $sheet->setCellValue("I3", $arData["moneys"]["arMonthsDate"]["middle"]["republica"]["people"]);
    $sheet->setCellValue("L3", $arData["moneys"]["arMonthsDate"]["middle"]["ilya"]["people"]);
    $sheet->setCellValue("D3", $arData["moneys"]["arMonthsDate"]["middle"]["playa"]["price"]);
    $sheet->setCellValue("G3", $arData["moneys"]["arMonthsDate"]["middle"]["cabeza"]["price"]);
    $sheet->setCellValue("J3", $arData["moneys"]["arMonthsDate"]["middle"]["republica"]["price"]);
    $sheet->setCellValue("M3", $arData["moneys"]["arMonthsDate"]["middle"]["ilya"]["price"]);


    $sheet->setCellValue("N3", $arData["arTargetCount"]["middle"]);
    $sheet->setCellValue("R3", $arData["arConversions"]["middle"]);
    $sheet->setCellValue("O3", $arData["moneys"]["arMonthsForTotalData"]["middle"]["all_people"]);
    $sheet->setCellValue("P3", $arData["moneys"]["arMonthsForTotalData"]["middle"]["all_price"]);
    $sheet->setCellValue("Q3", $arData["moneys"]["arMonthsForTotalData"]["middle"]["middle_check"]);
    $sheet->setCellValue("S3", $arData["moneys"]["arMonthsForTotalData"]["middle"]["fact_price"]);
    $sheet->setCellValue("T3", $arData["moneys"]["arMonthsForTotalData"]["middle"]["forecast"]);

    $i = 4;
    foreach ($this->arDate["list"] as $keyDate) {
      $sheet->setCellValue("A".$i, $keyDate);
      $sheet->setCellValue("C".$i, $arData["moneys"]["arMonthsDate"]["dates"][$keyDate]["playa"]["people"]);
      $sheet->setCellValue("D".$i, $arData["moneys"]["arMonthsDate"]["dates"][$keyDate]["playa"]["price"]);
      $sheet->setCellValue("F".$i, $arData["moneys"]["arMonthsDate"]["dates"][$keyDate]["cabeza"]["people"]);
      $sheet->setCellValue("G".$i, $arData["moneys"]["arMonthsDate"]["dates"][$keyDate]["cabeza"]["price"]);
      $sheet->setCellValue("I".$i, $arData["moneys"]["arMonthsDate"]["dates"][$keyDate]["republica"]["people"]);
      $sheet->setCellValue("J".$i, $arData["moneys"]["arMonthsDate"]["dates"][$keyDate]["republica"]["price"]);
      $sheet->setCellValue("L".$i, $arData["moneys"]["arMonthsDate"]["dates"][$keyDate]["ilya"]["people"]);
      $sheet->setCellValue("M".$i, $arData["moneys"]["arMonthsDate"]["dates"][$keyDate]["ilya"]["price"]);
      $sheet->setCellValue("H".$i, $arData["arTargetsForArea"]["republica"]["dates"][$keyDate]);
      $sheet->setCellValue("K".$i, $arData["arTargetsForArea"]["ilya"]["dates"][$keyDate]);

      // ИТОГО
      $sheet->setCellValue("N".$i, $arData["arTargetCount"]["dates"][$keyDate]);
      $sheet->setCellValue("O".$i, $arData["moneys"]["arMonthsForTotalData"]["dates"][$keyDate]["all_people"]);
      $sheet->setCellValue("P".$i, $arData["moneys"]["arMonthsForTotalData"]["dates"][$keyDate]["all_price"]);
      $sheet->setCellValue("Q".$i, $arData["moneys"]["arMonthsForTotalData"]["dates"][$keyDate]["middle_check"]);
      $sheet->setCellValue("R".$i, $arData["arConversions"]["dates"][$keyDate]);
      $sheet->setCellValue("S".$i, $arData["moneys"]["arMonthsForTotalData"]["dates"][$keyDate]["fact_price"]);
      $sheet->setCellValue("T".$i, $arData["moneys"]["arMonthsForTotalData"]["dates"][$keyDate]["forecast"]);

      $i++;
    }


    header("Expires: Mon, 1 Apr 1974 05:00:00 GMT");
    header("Last-Modified: ".gmdate("D,d M YH:i:s")." GMT");
    header("Cache-Control: no-cache, must-revalidate");
    header("Pragma: no-cache");
    header("Content-type: application/vnd.ms-excel");
    header("Content-Disposition: attachment; filename=records_".date("m_d_Y").".xls");
    $objWriter = new \PHPExcel_Writer_Excel5($obExcel);
    $objWriter->save('php://output');
  }


// end class
}
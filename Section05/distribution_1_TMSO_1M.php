
<?php
chdir(dirname(__FILE__));
date_default_timezone_set('Europe/Moscow');
require_once("config.php");
include("sesvars_TMSO.php");
require_once('Classes/PHPExcel.php');
require_once('Classes/PHPExcel/Writer/Excel2007.php');
?>
<?php
date_default_timezone_set('Europe/Moscow');

$graphcolor = "&bgcolor=0xF0ffff&bgcolorchart=0xdfedf3&fade1=ff6600&fade2=ff6314&colorbase=0xfff3b3&reverse=1";
$graphcolorstack = "&bgcolor=0xF0ffff&bgcolorchart=0xdfedf3&fade1=ff6600&colorbase=fff3b3&reverse=1&fade2=0x528252";

// ABANDONED CALLS

$query = "SELECT qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, ";
$query.= "qs.info1 AS info1, qs.info2 AS info2,  qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, ";
$query.= "qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND ";
$query.= "qs.qevent = ac.event_id AND qs.datetime >= '$start' AND qs.datetime <= '$end' ";
$query.= "AND q.queue IN ($queue,'NONE') AND ac.event IN ('ABANDON', 'EXITWITHTIMEOUT','COMPLETECALLER','COMPLETEAGENT','AGENTLOGIN','AGENTLOGOFF','AGENTCALLBACKLOGIN','AGENTCALLBACKLOGOFF') ";
$query.= "ORDER BY qs.datetime";

$query_comb     = "";
$login          = 0;
$logoff         = 0;
$dias           = Array();
$logout_by_day  = Array();
$logout_by_hour = Array();
$logout_by_dw   = Array();
$login_by_day   = Array();
$login_by_hour  = Array();
$login_by_dw    = Array();

$res = consulta_db($query,$DB_DEBUG,$DB_MUERE);

if(db_num_rows($res)>0) {

	while($row=db_fetch_row($res)) {
		$partes_fecha = split(" ",$row[0]);
		$partes_hora  = split(":",$partes_fecha[1]);

		$timestamp = return_timestamp($row[0]);
		$day_of_week = date('w',$timestamp);
			
		$dias[] = $partes_fecha[0];
		$horas[] = $partes_hora[0];

		if($row[3]=="ABANDON"){ //|| $row[3]=="EXITWITHTIMEOUT") {
		// проверяем не в ночное ли время не отвеченный
		$hour_call=date("G",strtotime($row[0]));
                echo "\n не принятый1 ".$row[3]." ".$row[2]." ".$row[1]." ".date("G",strtotime($row[0]))." \n";
			if ($hour_call>=6 and $hour_call <= 22 )
			{
				$unans_by_dw["$day_of_week"]++;
	               		echo "\n прошел1 ".$row[3]." ".$row[2]." ".$row[1]." ".date("G",strtotime($row[0]))." \n";
			}
                                $unanswered++;
                                $unans_by_day["$partes_fecha[0]"]++;
                                $unans_by_hour["$partes_hora[0]"]++;

		}	
		if($row[3]=="COMPLETECALLER" || $row[3]=="COMPLETEAGENT") {
 			$answered++;
			$ans_by_day["$partes_fecha[0]"]++;
			$ans_by_hour["$partes_hora[0]"]++;
			$ans_by_dw["$day_of_week"]++;

			$total_time_by_day["$partes_fecha[0]"]+=$row[5];
			$total_hold_by_day["$partes_fecha[0]"]+=$row[4];

			$total_time_by_dw["$day_of_week"]+=$row[5];
			$total_hold_by_dw["$day_of_week"]+=$row[4];
		
			$total_time_by_hour["$partes_hora[0]"]+=$row[5];
			$total_hold_by_hour["$partes_hora[0]"]+=$row[4];
		}
		if($row[3]=="AGENTLOGIN" || $row[3]=="AGENTCALLBACKLOGIN") {
 			$login++;
			$login_by_day["$partes_fecha[0]"]++;
			$login_by_hour["$partes_hora[0]"]++;
			$login_by_dw["$day_of_week"]++;
		}
		if($row[3]=="AGENTLOGOFF" || $row[3]=="AGENTCALLBACKLOGOFF") {
 			$logoff++;
			$logout_by_day["$partes_fecha[0]"]++;
			$logout_by_hour["$partes_hora[0]"]++;
			$logout_by_dw["$day_of_week"]++;
		}
	}
	$total_calls = $answered + $unanswered;
	$dias  = array_unique($dias);
	$horas = array_unique($horas);
    asort($dias);
	asort($horas);
} else {
 	// No rows returned
	$answered = 0;
	$unanswered = 0;
}
$start_parts = split(" ", $start);
$end_parts   = split(" ", $end);

$cover_pdf = $lang["$language"]['queue'].": ".$queue."\n";
$cover_pdf.= $lang["$language"]['start'].": ".$start_parts[0]."\n";
$cover_pdf.= $lang["$language"]['end'].": ".$end_parts[0]."\n";
$cover_pdf.= $lang["$language"]['period'].": ".$period." ".$lang["$language"]['days']."\n\n";
$cover_pdf.= $lang["$language"]['number_answered'].": ".$answered." ".$lang["$language"]['calls']."\n";
$cover_pdf.= $lang["$language"]['number_unanswered'].": ".$unanswered." ".$lang["$language"]['calls']."\n";
$cover_pdf.= $lang["$language"]['agent_login'].": ".$login."\n";
$cover_pdf.= $lang["$language"]['agent_logoff'].": ".$logoff."\n";
?>
<?php include("menu.php"); ?>

	
			<?php
				if(count($dias)<=0) {
					$dias['']=0;
				}
			?>

				<?php
				$header_pdf=array($lang["$language"]['date'],$lang["$language"]['answered'],$lang["$language"]['percent_answered'],$lang["$language"]['unanswered'],$lang["$language"]['percent_unanswered'],$lang["$language"]['avg_calltime'],$lang["$language"]['avg_holdtime'],$lang["$language"]['login'],$lang["$language"]['logoff']);
				$width_pdf=array(25,23,23,23,23,25,25,20,20);
				$title_pdf=$lang["$language"]['call_distrib_day'];

				$count=1;
				foreach($dias as $key) {
					$cual = $count%2;
					if($cual>0) { $odd = " class='odd' "; } else { $odd = ""; }
					if(!isset($ans_by_day["$key"])) {
						$ans_by_day["$key"]=0;
					}
					if(!isset($unans_by_day["$key"])) {
						$unans_by_day["$key"]=0;
					}
					if($answered > 0) {
						$percent_ans   = $ans_by_day["$key"]   * 100 / $answered;
					} else {
						$percent_ans = 0;
					}
					if($ans_by_day["$key"] >0) {
						$average_call_duration = $total_time_by_day["$key"] / $ans_by_day["$key"];
						$average_hold_duration = $total_hold_by_day["$key"] / $ans_by_day["$key"];
					} else {
						$average_call_duration = 0;
						$average_hold_duration = 0;
					}
					if($unanswered > 0) {
						$percent_unans = $unans_by_day["$key"] * 100 / $unanswered;
					} else {
						$percent_unans = 0;
					}
					$percent_ans   = number_format($percent_ans,  2);
					$percent_unans = number_format($percent_unans,2);
					$average_call_duration_print = seconds2minutes($average_call_duration);
					if($key<>"") {
					$linea_pdf = array($key,$ans_by_day["$key"],"$percent_ans ".$lang["$language"]['percent'],$unans_by_day["$key"],"$percent_unans ".$lang["$language"]['percent'],$average_call_duration_print,number_format($average_hold_duration,0),$login_by_day["$key"],$logout_by_day["$key"]);
					$data_pdf[]=$linea_pdf;				
					
					}
				}
				
				?>
				<?php

				$header_pdf=array($lang["$language"]['hour'],$lang["$language"]['answered'],$lang["$language"]['percent_answered'],$lang["$language"]['unanswered'],$lang["$language"]['percent_unanswered'],$lang["$language"]['avg_calltime'],$lang["$language"]['avg_holdtime'],$lang["$language"]['login'],$lang["$language"]['logoff']);
				$width_pdf=array(25,23,23,23,23,25,25,20,20);
				$title_pdf=$lang["$language"]['call_distrib_hour'];
				$data_pdf = array();

				$query_ans = "";
				$query_unans = "";
				$query_time="";
				$query_hold="";
				for($key=0;$key<24;$key++) {
					$cual = ($key+1)%2;
					if($cual>0) { $odd = " class='odd' "; } else { $odd = ""; }
					if(strlen($key)==1) { $key = "0".$key; }
					if(!isset($ans_by_hour["$key"])) {
						$ans_by_hour["$key"]=0;
						$average_call_duration = 0;
						$average_hold_duration = 0;
					} else {
						$average_call_duration = $total_time_by_hour["$key"] / $ans_by_hour["$key"];
						$average_hold_duration = $total_hold_by_hour["$key"] / $ans_by_hour["$key"];
					}
					if(!isset($unans_by_hour["$key"])) {
						$unans_by_hour["$key"]=0;
					}
					if($answered > 0) {
						$percent_ans   = $ans_by_hour["$key"]   * 100 / $answered;
					} else {
						$percent_ans = 0;
					}
					if($unanswered > 0) {
						$percent_unans = $unans_by_hour["$key"] * 100 / $unanswered;
					} else {
						$percent_unans = 0;
					}
					$percent_ans   = number_format($percent_ans,  2);
					$percent_unans = number_format($percent_unans,2);

					if(!isset($login_by_hour["$key"])) {
					    $login_by_hour["$key"]=0;
                    }
					if(!isset($logout_by_hour["$key"])) {
					    $logout_by_hour["$key"]=0;
                    }

					$linea_pdf = array($key,$ans_by_hour["$key"],"$percent_ans ".$lang["$language"]['percent'],$unans_by_hour["$key"],"$percent_unans ".$lang["$language"]['percent'],number_format($average_call_duration,0),number_format($average_hold_duration,0),$login_by_hour["$key"],$logout_by_hour["$key"]);

					$gkey = $key+1;
					$query_ans  .="var$gkey=$key&val$gkey=".$ans_by_hour["$key"]."&";
					$query_unans.="var$gkey=$key&val$gkey=".$unans_by_hour["$key"]."&";
					$query_comb.= "var$gkey=$key%20".$lang["$language"]['hours']."&valA$gkey=".$ans_by_hour["$key"]."&valB$gkey=".$unans_by_hour["$key"]."&";
					$query_time.="var$gkey=$key&val$gkey=".intval($average_call_duration)."&";
					$query_hold.="var$gkey=$key&val$gkey=".intval($average_hold_duration)."&";
					$data_pdf[]=$linea_pdf;
				
				}
				
				
			//мой код
			   $dayinyear = (date('z',$timestamp));
			   //$numberrpt = str_pad($dayinyear, 4, '0', STR_PAD_LEFT);
			   $numberrpt = sprintf('%04d', $dayinyear);	
			   $objPHPExcel = new PHPExcel(); //Объект PHPExcel
               $objPHPExcel = PHPExcel_IOFactory::createReader('Excel2007'); //Задаем ридер
			   $objPHPExcel->setIncludeCharts(TRUE);
               $objPHPExcel = $objPHPExcel->load('DayReport_TMSO_3.xlsx'); //Загружаем "шаблонный" xls

// выставляем дату "жестко"

                $objPHPExcel->getSheetByName('Report')
                        ->setCellValue('J3', date("d.m.Y", strtotime($start)))
                ;


               $objWriter =  new PHPExcel_Writer_Excel2007($objPHPExcel);			  

               $name='UTS-M11-S00-ITM-RPT-'.$numberrpt.'-'.date("Y-m-d",time()-86400).'-M11 15-58 TMCO call Statistics'.'.xlsx';
			   $objWriter->setIncludeCharts(true);
               $objWriter->save(dirname(__FILE__)."/dir1"."/".$name);
               $objPHPexcel = PHPExcel_IOFactory::load(dirname(__FILE__)."/dir1"."/".$name);   
			   $objPHPExcel->setActiveSheetIndex(1);

			$objPHPExcel->getActiveSheet()->fromArray($header_pdf, null, 'A1');			   
			   $objPHPExcel->getActiveSheet()->fromArray($header_pdf, null, 'A1');
			   $objPHPExcel->getActiveSheet()->fromArray($data_pdf, null, 'A2');
					
				//конец кода
	
		
				$query_ans.="title=".$lang["$language"]['answ_by_hour']."$graphcolor";
				$query_unans.="title=".$lang["$language"]['unansw_by_hour']."$graphcolor";
				$query_time.="title=".$lang["$language"]['avg_call_time_by_hr']."$graphcolor";
				$query_hold.="title=".$lang["$language"]['avg_hold_time_by_hr']."$graphcolor";
				$query_comb.="title=".$lang["$language"]['anws_unanws_by_hour']."$graphcolorstack&tagA=".$lang["$language"]['answered_calls']."&tagB=".$lang["$language"]['unanswered_calls'];
				?>

	

				<?php
				//Заново загружаем данные в таблицы
				
//Откатиться на начало недели
$day_of_week = date('w',$timestamp);
$delta=86400*($day_of_week+1);
$start = date('Y-m-d 00:00:00',time()-$delta);
$end = date('Y-m-d 00:00:00',time()-86400);
//Конец отката

$query = "SELECT qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, ";
$query.= "qs.info1 AS info1, qs.info2 AS info2,  qs.info3 AS info3 FROM queue_stats AS qs, qname AS q, ";
$query.= "qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND ";
$query.= "qs.qevent = ac.event_id AND qs.datetime >= '$start' AND qs.datetime <= '$end' ";
$query.= "AND q.queue IN ($queue,'NONE') AND ac.event IN ('ABANDON', 'EXITWITHTIMEOUT','COMPLETECALLER','COMPLETEAGENT','AGENTLOGIN','AGENTLOGOFF','AGENTCALLBACKLOGIN','AGENTCALLBACKLOGOFF') ";
$query.= "ORDER BY qs.datetime";
$res = consulta_db($query,$DB_DEBUG,$DB_MUERE);

//Обнуление всего перед циклом
$answered = 0;
$unanswered = 0;





if(db_num_rows($res)>0) {

	while($row=db_fetch_row($res)) {
		$partes_fecha = split(" ",$row[0]);
		$partes_hora  = split(":",$partes_fecha[1]);

		$timestamp = return_timestamp($row[0]);
		$day_of_week = date('w',$timestamp);
			
		$dias[] = $partes_fecha[0];
		$horas[] = $partes_hora[0];

		if($row[3]=="ABANDON"){ // || $row[3]=="EXITWITHTIMEOUT") {
			$hour_call=date("G",strtotime($row[0]));
			echo "\n не принятый2 ".$row[3]." ".$row[2]." ".$row[1]." ".date("G",strtotime($row[0]))." ".$row[0]." \n";
                        if ($hour_call>=6 and $hour_call <= 22 )
                        {
                                $unans_by_dw["$day_of_week"]++;
                                echo "\n прошел2 ".$row[3]." ".$row[2]." ".$row[1]." ".date("G",strtotime($row[0]))." ".$row[0]." \n";
                        }
			$unanswered++;
			$unans_by_day["$partes_fecha[0]"]++;
			$unans_by_hour["$partes_hora[0]"]++;
		}
		if($row[3]=="COMPLETECALLER" || $row[3]=="COMPLETEAGENT") {
 			$answered++;
			$ans_by_day["$partes_fecha[0]"]++;
			$ans_by_hour["$partes_hora[0]"]++;
			$ans_by_dw["$day_of_week"]++;

			$total_time_by_day["$partes_fecha[0]"]+=$row[5];
			$total_hold_by_day["$partes_fecha[0]"]+=$row[4];

			$total_time_by_dw["$day_of_week"]+=$row[5];
			$total_hold_by_dw["$day_of_week"]+=$row[4];
		
			$total_time_by_hour["$partes_hora[0]"]+=$row[5];
			$total_hold_by_hour["$partes_hora[0]"]+=$row[4];
		}
		if($row[3]=="AGENTLOGIN" || $row[3]=="AGENTCALLBACKLOGIN") {
 			$login++;
			$login_by_day["$partes_fecha[0]"]++;
			$login_by_hour["$partes_hora[0]"]++;
			$login_by_dw["$day_of_week"]++;
		}
		if($row[3]=="AGENTLOGOFF" || $row[3]=="AGENTCALLBACKLOGOFF") {
 			$logoff++;
			$logout_by_day["$partes_fecha[0]"]++;
			$logout_by_hour["$partes_hora[0]"]++;
			$logout_by_dw["$day_of_week"]++;
		}
	}
	$total_calls = $answered + $unanswered;
	$dias  = array_unique($dias);
	$horas = array_unique($horas);
    asort($dias);
	asort($horas);
} else {
 	// No rows returned
	$answered = 0;
	$unanswered = 0;
}
				
	
				$header_pdf=array($lang["$language"]['day'],$lang["$language"]['answered'],$lang["$language"]['percent_answered'],$lang["$language"]['unanswered'],$lang["$language"]['percent_unanswered'],$lang["$language"]['avg_calltime'],$lang["$language"]['avg_holdtime'],$lang["$language"]['login'],$lang["$language"]['logoff']);
				$width_pdf=array(25,23,23,23,23,25,25,20,20);
				$title_pdf=$lang["$language"]['call_distrib_week'];
				$data_pdf = array();


				$query_ans="";
				$query_unans="";
				$query_time="";
				$query_hold="";
				for($key=0;$key<7;$key++) {
					$cual = ($key+1)%2;
					if($cual>0) { $odd = " class='odd' "; } else { $odd = ""; }
					if(!isset($total_time_by_dw["$key"])) {
						$total_time_by_dw["$key"]=0;
					}
					if(!isset($total_hold_by_dw["$key"])) {
						$total_hold_by_dw["$key"]=0;
					}
					if(!isset($ans_by_dw["$key"])) {
						$ans_by_dw["$key"]=0;
						$average_call_duration = 0;
						$average_hold_duration = 0;
					} else {
						$average_call_duration = $total_time_by_dw["$key"] / $ans_by_dw["$key"];
						$average_hold_duration = $total_hold_by_dw["$key"] / $ans_by_dw["$key"];
					}

					if(!isset($unans_by_dw["$key"])) {
						$unans_by_dw["$key"]=0;
					}
					if($answered > 0) {
						$percent_ans   = $ans_by_dw["$key"]   * 100 / $answered;
					} else {
						$percent_ans = 0;
					}
					if($unanswered > 0) {
						$percent_unans = $unans_by_dw["$key"] * 100 / $unanswered;
					} else {
						$percent_unans = 0;
					}
					$percent_ans   = number_format($percent_ans,  2);
					$percent_unans = number_format($percent_unans,2);

					if(!isset($login_by_dw["$key"])) {
					    $login_by_dw["$key"]=0;
                    }
					if(!isset($logout_by_dw["$key"])) {
					    $logout_by_dw["$key"]=0;
                    }

					$linea_pdf = array($dayp["$key"],$ans_by_dw["$key"],"$percent_ans ".$lang["$language"]['percent'],$unans_by_dw["$key"],"$percent_unans ".$lang["$language"]['percent'],number_format($average_call_duration,0),number_format($average_hold_duration,0),$login_by_dw["$key"],$logout_by_dw["$key"]);

					$gkey = $key + 1;
					$query_ans  .="var$gkey=".$dayp["$key"]."&val$gkey=".intval($ans_by_dw["$key"])."&";
					$query_unans.="var$gkey=".$dayp["$key"]."&val$gkey=".intval($unans_by_dw["$key"])."&";
					$query_time.="var$gkey=".$dayp["$key"]."&val$gkey=".intval($average_call_duration)."&";
					$query_hold.="var$gkey=".$dayp["$key"]."&val$gkey=".intval($average_hold_duration)."&";
					$data_pdf[]=$linea_pdf;
				}
				$query_ans.="title=".$lang["$language"]['answ_by_day']."$graphcolor";
				$query_unans.="title=".$lang["$language"]['unansw_by_day']."$graphcolor";
				$query_time.="title=".$lang["$language"]['avg_call_time_by_day']."$graphcolor";
				$query_hold.="title=".$lang["$language"]['avg_hold_time_by_day']."$graphcolor";
				
				
				//Добавление на 3 лист				
			   $objPHPExcel->setActiveSheetIndex(3);			   
			//   $objPHPExcel->getActiveSheet()->fromArray($header_pdf, null, 'A1');
			   $objPHPExcel->getActiveSheet()->fromArray($data_pdf, null, 'A1');
			   
				
				//Конец добавления распределения по неделе
				
				
				
				
				
				?>
<?php
//Дату обратно

$start = date('Y-m-d 00:00:00',time()-86400);
$end = date('Y-m-d 23:59:00',time()-86400);
//Конец

$query_comb     = "";
$login          = 0;
$logoff         = 0;
$dias           = Array();
$logout_by_day  = Array();
$logout_by_hour = Array();
$logout_by_dw   = Array();
$login_by_day   = Array();
$login_by_hour  = Array();
$login_by_dw    = Array();

$graphcolor2 = "&bgcolor=0xF0ffff&bgcolorchart=0xdfedf3&fade1=0xff6600&fade2=0x528252&colorbase=0xfff3b3&reverse=1";
$graphcolor  = "&bgcolor=0xF0ffff&bgcolorchart=0xdfedf3&fade1=0xff6600&fade2=0xff6600&colorbase=0xfff3b3&reverse=1";
// This query shows the hangup cause, how many calls an
// agent hanged up, and a caller hanged up.
$query = "SELECT count(ev.event) AS num, ev.event AS action ";
$query.= "FROM queue_stats AS qs, qname AS q, qevent AS ev WHERE ";
$query.= "qs.qname = q.qname_id and qs.qevent = ev.event_id and qs.datetime >= '$start' and ";
$query.= "qs.datetime <= '$end' and q.queue IN ($queue) AND ";
$query.= "ev.event IN ('COMPLETECALLER', 'COMPLETEAGENT') ";
$query.= "GROUP BY ev.event ORDER BY ev.event";

$hangup_cause["COMPLETECALLER"]=0;
$hangup_cause["COMPLETEAGENT"]=0;
$res = consulta_db($query,$DB_DEBUG,$DB_MUERE);
while($row=db_fetch_row($res)) {
  $hangup_cause["$row[1]"]=$row[0];
  $total_hangup+=$row[0];
}



$query = "SELECT qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, "; 
$query.= "ac.event AS qevent, qs.info1 AS info1, qs.info2 AS info2,  qs.info3 AS info3 ";
$query.= "FROM queue_stats AS qs, qname AS q, qagent AS ag, qevent AS ac WHERE ";
$query.= "qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND qs.qevent = ac.event_id AND ";
$query.= "qs.datetime >= '$start' AND qs.datetime <= '$end' AND ";
$query.= "q.queue IN ($queue) AND ag.agent in ($agent) AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT','TRANSFER','CONNECT') ORDER BY qs.datetime";

$answer["15"]=0;
$answer["30"]=0;
$answer["45"]=0;
$answer["60"]=0;
$answer["75"]=0;
$answer["90"]=0;
$answer["91+"]=0;

$abandoned         = 0;
$transferidas      = 0;
$totaltransfers    = 0;
$total_hangup      = 0;
$total_calls       = 0;
$total_calls2      = Array();
$total_duration    = 0;
$total_calls_queue = Array();

$res = consulta_db($query,$DB_DEBUG,$DB_MUERE);
if($res) {
    while($row=db_fetch_row($res)) {
        if($row[3] <> "TRANSFER" && $row[3]<>"CONNECT") {
            $total_hold     += $row[4];
            $total_duration += $row[5];
            $total_calls++;
            $total_calls_queue["$row[1]"]++;
        } elseif($row[3]=="TRANSFER") {
            $transferidas++;
        }
        if($row[3]=="CONNECT") {

            if ($row[4] >=0 && $row[4] <= 15) {
                $answer["15"]++;
            }

            if ($row[4] >=16 && $row[4] <= 30) {
                $answer["30"]++;
            }

            if ($row[4] >=31 && $row[4] <= 45) {
              $answer["45"]++;
            }

            if ($row[4] >=46 && $row[4] <= 60) {
              $answer["60"]++;
            }

            if ($row[4] >=61 && $row[4] <= 75) {
              $answer["75"]++;
            }

            if ($row[4] >=76 && $row[4] <= 90) {
              $answer["90"]++;
            }

            if ($row[4] >=91) {
              $answer["91+"]++;
            }
        }
    }
} 

if($total_calls > 0) {
    ksort($answer);
    $average_hold     = $total_hold     / $total_calls;
    $average_duration = $total_duration / $total_calls;
    $average_hold     = number_format($average_hold     , 2);
    $average_duration = number_format($average_duration , 2);
} else {
    // There were no calls
    $average_hold = 0;
    $average_duration = 0;
}

$total_duration_print = seconds2minutes($total_duration);
// TRANSFERS
$query = "SELECT ag.agent AS agent, qs.info1 AS info1,  qs.info2 AS info2 ";
$query.= "FROM  queue_stats AS qs, qevent AS ac, qagent as ag, qname As q WHERE qs.qevent = ac.event_id ";
$query.= "AND qs.qname = q.qname_id AND ag.agent_id = qs.qagent AND qs.datetime >= '$start' ";
$query.= "AND qs.datetime <= '$end' AND  q.queue IN ($queue)  AND ag.agent in ($agent) AND  ac.event = 'TRANSFER'";


$res = consulta_db($query,$DB_DEBUG,$DB_MUERE);
if($res) {
    while($row=db_fetch_row($res)) {
        $keytra = "$row[0]^$row[1]@$row[2]";
        $transfers["$keytra"]++;
        $totaltransfers++;
    }
} else {
   $totaltransfers=0;
}

// ABANDONED CALLS
$query = "SELECT  ac.event AS action,  qs.info1 AS info1,  qs.info2 AS info2,  qs.info3 AS info3";
$query.= "FROM  queue_stats AS qs, qevent AS ac, qname As q, qagent as ag WHERE ";
$query.= "qs.qevent = ac.event_id AND qs.qname = q.qname_id AND qs.datetime >= '$start' AND ";
$query.= "qs.datetime <= '$end' AND  q.queue IN ($queue)  AND ag.agent in ($agent) AND  ac.event IN ('ABANDON', 'EXITWITHTIMEOUT', 'TRANSFER') ";
$query.= "ORDER BY  ac.event,  qs.info3";

$res = consulta_db($query,$DB_DEBUG,$DB_MUERE);

while($row=db_fetch_row($res)) {

    if($row[0]=="ABANDON") {
	echo "\n не принятый3 ".$row[3]." ".$row[2]." ".$row[1]." ".$row[0]." ".date("G",strtotime($row[0]))." \n";
         $abandoned++;
        $abandon_end_pos+=$row[1];
        $abandon_start_pos+=$row[2];
        $total_hold_abandon+=$row[3];
    }
    if($row[0]=="EXITWITHTIMEOUT") {
         $timeout++;
        $timeout_end_pos+=$row[1];
        $timeout_start_pos+=$row[2];
        $total_hold_timeout+=$row[3];
    }
}

if($abandoned > 0) {
    $abandon_average_hold = $total_hold_abandon / $abandoned;
    $abandon_average_hold = number_format($abandon_average_hold,2);

    $abandon_average_start = floor($abandon_start_pos / $abandoned);
    $abandon_average_start = number_format($abandon_average_start,2);

    $abandon_average_end = floor($abandon_end_pos / $abandoned);
    $abandon_average_end = number_format($abandon_average_end,2);
} else {
    $abandoned = 0;
    $abandon_average_hold  = 0;
    $abandon_average_start = 0;
    $abandon_average_end   = 0;
}

// This query shows every call for agents, we collect into a named array the values of holdtime and calltime
$query = "SELECT qs.datetime AS datetime, q.queue AS qname, ag.agent AS qagent, ac.event AS qevent, ";
$query.= "qs.info1 AS info1, qs.info2 AS info2, qs.info3 AS info3  FROM queue_stats AS qs, qname AS q, ";
$query.= "qagent AS ag, qevent AS ac WHERE qs.qname = q.qname_id AND qs.qagent = ag.agent_id AND ";
$query.= "qs.qevent = ac.event_id AND qs.datetime >= '$start' AND qs.datetime <= '$end' AND ";
$query.= "q.queue IN ($queue) AND ag.agent in ($agent) AND ac.event IN ('COMPLETECALLER', 'COMPLETEAGENT') ORDER BY ag.agent";

$res = consulta_db($query,$DB_DEBUG,$DB_MUERE);


//попробую результат на лист эксель

// Мои переменные по операторам

		$dias[] = $partes_fecha[0];
		$horas[] = $partes_hora[0];


	while($row=db_fetch_row($res)) {
		$partes_fecha = split(" ",$row[0]);
		$partes_hora  = split(":",$partes_fecha[1]);

		$timestamp = return_timestamp($row[0]);
		$day_of_week = date('w',$timestamp);
			
		$dias[] = $partes_fecha[0];
		$horas[] = $partes_hora[0];


//3243 ТМСО1 ряд 1
		if(strpos($row[2],"TMCO 1") !== FALSE) {
 			$answered21++;
			$ans_by_day21["$partes_fecha[0]"]++;
			$ans_by_hour21["$partes_hora[0]"]++;
			$ans_by_dw21["$day_of_week"]++;
		}

//3244 ТМСО2 ряд 2
		if(strpos($row[2],"TMCO 2")!== FALSE) {
 			$answered22++;
			$ans_by_day22["$partes_fecha[0]"]++;
			$ans_by_hour22["$partes_hora[0]"]++;
			$ans_by_dw22["$day_of_week"]++;
		}



//3945 ряд 5
                 if(strpos($row[2],"3945") !== FALSE){
                        $answeredt22++;
                        $ans_by_dayt22["$partes_fecha[0]"]++;
                        $ans_by_hourt22["$partes_hora[0]"]++;
                        $ans_by_dwt22["$day_of_week"]++;
                }


//3245 ТМСО 3 ряд 3
		if(strpos($row[2],"TMCO 3") !== FALSE) {
 			$answered23++;
			$ans_by_day23["$partes_fecha[0]"]++;
			$ans_by_hour23["$partes_hora[0]"]++;
			$ans_by_dw23["$day_of_week"]++;
		}



// 3944 ряд 4		
		if(strpos($row[2],"3944") !== FALSE) {
 			$answered24++;
			$ans_by_day24["$partes_fecha[0]"]++;
			$ans_by_hour24["$partes_hora[0]"]++;
			$ans_by_dw24["$day_of_week"]++;
		}

/*

		if(strpos($row[2],"TMCO 4") !== FALSE) {
			
 			$answered25++;
			$ans_by_day25["$partes_fecha[0]"]++;
			$ans_by_hour25["$partes_hora[0]"]++;
			$ans_by_dw25["$day_of_week"]++;
				
		}
*/

	}
	$dias  = array_unique($dias);
	$horas = array_unique($horas);
    asort($dias);
	asort($horas);

//попробуем заполнить по операторам
	$data_pdf = array();
			for($key=0;$key<24;$key++) {
					$cual = ($key+1)%2;
					if($cual>0) { $odd = " class='odd' "; } else { $odd = ""; }
					if(strlen($key)==1) { $key = "0".$key; }
					if(!isset($ans_by_hour21["$key"])) {
						$ans_by_hour21["$key"]=0;
					} 
					if(!isset($ans_by_hour22["$key"])) {
						$ans_by_hour22["$key"]=0;
					} 
					if(!isset($ans_by_hour23["$key"])) {
						$ans_by_hour23["$key"]=0;
					} 
					if(!isset($ans_by_hour24["$key"])) {
						$ans_by_hour24["$key"]=0;
					} 			
					if(!isset($ans_by_hourt22["$key"])) {
                                                $ans_by_hourt22["$key"]=0;
                                        }
					$linea_pdf = array($key,$ans_by_hour21["$key"],$ans_by_hour22["$key"],$ans_by_hour23["$key"],$ans_by_hour24["$key"],$ans_by_hourt22["$key"],$ans_by_hour25["$key"]);
					$gkey = $key+1;
					$data_pdf[]=$linea_pdf;
				}





//$new_array[] = $row;
//while ($row = db_fetch_row($res)) {
	
    //$partes_fecha = split(" ",$row[0]);
	//$partes_hora  = split(":",$partes_fecha[1]);	
	
	
    //$new_array[$row['datetime']] = $partes_hora[1];
    //$new_array[$row['qname']] = $row;
	//$new_array[$row['qagent']] = $row;
	//$new_array[$row['qevent']] = $row;
	//$new_array[$row['info1']] = $row;
	//$new_array[$row['info2']] = $row;
	//$new_array[$row['info3']] = $row;	
//}







$objPHPExcel->setActiveSheetIndex(4);
$objPHPExcel->getActiveSheet()->fromArray($data_pdf, null, 'A1');



//попробовал

while($row=db_fetch_row($res)) {
    $total_calls2["$row[2]"]++;
    $record["$row[2]"][]=$row[0]."|".$row[1]."|".$row[3]."|".$row[4];
    $total_hold2["$row[2]"]+=$row[4];
    $total_time2["$row[2]"]+=$row[5];
    $grandtotal_hold+=$row[4];
    $grandtotal_time+=$row[5];
    $grandtotal_calls++;
}

$start_parts = split(" ", $start);
$end_parts   = split(" ", $end);

$cover_pdf = $lang["$language"]['queue'].": ".$queue."\n";
$cover_pdf.= $lang["$language"]['start'].": ".$start_parts[0]."\n";
$cover_pdf.= $lang["$language"]['end'].": ".$end_parts[0]."\n";
$cover_pdf.= $lang["$language"]['period'].": ".$period." ".$lang["$language"]['days']."\n\n";
$cover_pdf.= $lang["$language"]['answered_calls'].": ".$total_calls." ".$lang["$language"]['calls']."\n";
$cover_pdf.= $lang["$language"]['avg_calltime'].": ".$average_duration." ".$lang["$language"]['secs']."\n";
$cover_pdf.= $lang["$language"]['total'].": ".$total_duration_print." ".$lang["$language"]['minutes']."\n";
$cover_pdf.= $lang["$language"]['avg_holdtime'].": ".$average_hold." ".$lang["$language"]['secs']."\n";

?>

                <?php
				// Второй лист
				$objPHPExcel->setActiveSheetIndex(2);	
				$countstr="";
                $countrow=0;
                $partial_total = 0;
                $query2="";
                $total_y_transfer = $answer['15'] + $answer['30'] +  $answer['45'] + $answer['60'] +  $answer['75'] + $answer['90'] +  $answer['91+'];
                foreach($answer as $key=>$val)
                {
                    $newcont = $countrow+1;
                    $query2.="val$newcont=$val&var$newcont=$key%20".$lang["$language"]['secs']."&";
                    $cual = ($countrow%2);
                    if($cual>0) { $odd = " class='odd' "; } else { $odd = ""; }                  
                    $delta = $val;
                    if($delta > 0) { $delta = "+".$delta;}
                    $partial_total += $val;
                    if($total_y_transfer > 0) {
                    $percent=$partial_total*100/$total_y_transfer;
                    } else {
                    $percent = 0;
                    }
                    $percent=number_format($percent,2);
                    if($countrow==0) { $delta = ""; }                    
					//Загоняем в таблицу
					$countstr="B".strval($countrow+2);
					$objPHPExcel->getActiveSheet()->setCellValue($countstr,$partial_total);
                    $countrow++;
                }
                $query2.="title=".$lang["$language"]['call_response']."$graphcolor";
                ?>			

			    <?php
				 //Заливка второй таблицы упаковка и отсылка
			   
			   //$objPHPExcel->getActiveSheet()->fromArray($header_pdf, null, 'A1');
			   $objPHPExcel->setActiveSheetIndex(0);			   


// собираю статистику по src (callerid)

/*
$query = "select count(*), q.*, n.*, e.*, a.* , aa.description from queue_stats q left join qname n on q.qname=n.qname_id left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id ";
$query.= "left join asterisk.incoming aa on q.info2=aa.cidnum where q.datetime>=$start and q.datetime<=$end and n.queue=9995 and (e.event='ENTERQUEUE') ";
$query.= "group by q.info2 order by count(*) desc limit 10";
$res = consulta_db($query,$DB_DEBUG,$DB_MUERE);
*/

$query1 = "select  aa.description as 'description', count(*) as 'counts', q.info2 as 'callerid' from queue_stats q left join qname n on q.qname=n.qname_id ";
$query1.= "left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id left join asterisk.incoming aa on q.info2=aa.cidnum where q.datetime>='$start' and q.datetime<='$end' ";
$query1.= "and n.queue=9995 and (e.event='ENTERQUEUE') group by q.info2 order by count(*) desc limit 5";

$res1 = consulta_db($query1,$DB_DEBUG,$DB_MUERE);

echo "\n  ".$query1;

//записываю статистику по src на лист 4
$count_i=0;
$objPHPExcel->setActiveSheetIndex(4);

while($row=db_fetch_row($res1)) 
{
    $result_q1['description'][$count_i]=$row[0];
    $result_q1['counts'][$count_i]=$row[1];
    $result_q1['callerid'][$count_i]=$row[2];
// хардкол для CC APD
    if ($row[2]=='74952499390')
    {
	$result_q1['description'][$count_i]="APD CC";
    }
    $count_i++;
}

//print_r($result_q1);

// запрос общео и потом по категориям

$query2 = "select count(*) as 'counts' from queue_stats q left join qname n on q.qname=n.qname_id left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id where q.datetime>='$start' " ;
$query2.= "and q.datetime<='$end' and n.queue=9995 and (e.event='COMPLETECALLER' or e.event='COMPLETEAGENT' or e.event='ABANDON')";
$res1 = consulta_db($query2,$DB_DEBUG,$DB_MUERE);

while($row=db_fetch_row($res1))
{
    $result_q1['description'][$count_i]="Other";
    $result_q1['counts'][$count_i]=$row[0];
    $result_q1['callerid'][$count_i]=' ';
}
$total_inc1=$result_q1['counts'][$count_i];
$result_q1['counts'][$count_i]=$result_q1['counts'][$count_i]-($result_q1['counts'][0]+$result_q1['counts'][1]+$result_q1['counts'][2]+$result_q1['counts'][3]+$result_q1['counts'][4]);
print_r($result_q1);
$count_i++;


//Запросы по категориям source

// TSI
$count_i=0;
$query3 = "select count(*) as 'counts' from "; 
$query3.= "queue_stats q left join qname n on q.qname=n.qname_id left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id left join asterisk.incoming aa on q.info2=aa.cidnum "; 
$query3.= "where q.datetime>='$start' and q.datetime<='$end' and n.queue=9995 and (e.event='ENTERQUEUE') and aa.description like '%TSIS%'";
$res3 = consulta_db($query3,$DB_DEBUG,$DB_MUERE);

$result_q3['description'][$count_i]="TSI";
$result_q3['counts'][$count_i]=0;
while($row=db_fetch_row($res3))
{
    $result_q3['counts'][$count_i]=$row[0];
}

//GIBDD
$count_i++;
$query3 = "select count(*) as 'counts' from ";
$query3.= "queue_stats q left join qname n on q.qname=n.qname_id left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id left join asterisk.incoming aa on q.info2=aa.cidnum ";
$query3.= "where q.datetime>='$start' and q.datetime<='$end' and n.queue=9995 and (e.event='ENTERQUEUE') and aa.description like '%GIBDD%'";
$res3 = consulta_db($query3,$DB_DEBUG,$DB_MUERE);

$result_q3['description'][$count_i]="GIBDD";
$result_q3['counts'][$count_i]=0;
while($row=db_fetch_row($res3))
{
    $result_q3['counts'][$count_i]=$row[0];
}

//MTTS
$count_i++;
$query3 = "select count(*) as 'counts' from ";
$query3.= "queue_stats q left join qname n on q.qname=n.qname_id left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id left join asterisk.incoming aa on q.info2=aa.cidnum ";
$query3.= "where q.datetime>='$start' and q.datetime<='$end' and n.queue=9995 and (e.event='ENTERQUEUE') and aa.description like '%MTTS%'";
$res3 = consulta_db($query3,$DB_DEBUG,$DB_MUERE);

$result_q3['description'][$count_i]="MTTS";
$result_q3['counts'][$count_i]=0;
while($row=db_fetch_row($res3))
{
    $result_q3['counts'][$count_i]=$row[0];
}

//Evacuation
$count_i++;
$query3 = "select count(*) as 'counts' from ";
$query3.= "queue_stats q left join qname n on q.qname=n.qname_id left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id left join asterisk.incoming aa on q.info2=aa.cidnum ";
$query3.= "where q.datetime>='$start' and q.datetime<='$end' and n.queue=9995 and (e.event='ENTERQUEUE') and aa.description like '%Evacuation%'";
$res3 = consulta_db($query3,$DB_DEBUG,$DB_MUERE);

$result_q3['description'][$count_i]="Evacuation";
$result_q3['counts'][$count_i]=0;
while($row=db_fetch_row($res3))
{
    $result_q3['counts'][$count_i]=$row[0];
}

//RAS
$count_i++;
$query3 = "select count(*) as 'counts' from ";
$query3.= "queue_stats q left join qname n on q.qname=n.qname_id left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id left join asterisk.incoming aa on q.info2=aa.cidnum ";
$query3.= "where q.datetime>='$start' and q.datetime<='$end' and n.queue=9995 and (e.event='ENTERQUEUE') and aa.description like '%RAS%'";
$res3 = consulta_db($query3,$DB_DEBUG,$DB_MUERE);

$result_q3['description'][$count_i]="RA";
$result_q3['counts'][$count_i]=0;
while($row=db_fetch_row($res3))
{
    $result_q3['counts'][$count_i]=$row[0];
}

//APD CC

$count_i++;
$query3 = "select count(*) as 'counts' from ";
$query3.= "queue_stats q left join qname n on q.qname=n.qname_id left join qevent e on q.qevent=e.event_id left join qagent a on q.qagent=a.agent_id left join asterisk.incoming aa on q.info2=aa.cidnum ";
$query3.= "where q.datetime>='$start' and q.datetime<='$end' and n.queue=9995 and (e.event='ENTERQUEUE') and q.info2 in ('74952499390','74952491625')";
$res3 = consulta_db($query3,$DB_DEBUG,$DB_MUERE);

$result_q3['description'][$count_i]="APD CC";
$result_q3['counts'][$count_i]=0;
while($row=db_fetch_row($res3))
{
    $result_q3['counts'][$count_i]=$row[0];
}

//Other
$count_i++;
$result_q3['description'][$count_i]="Other";
$result_q3['counts'][$count_i]=$total_inc1-($result_q3['counts'][0]+$result_q3['counts'][1]+$result_q3['counts'][2]+$result_q3['counts'][3]+$result_q3['counts'][4]+$result_q3['counts'][5]);


$objPHPExcel->getActiveSheet()->fromArray($result_q1, null, 'A60');
$objPHPExcel->getActiveSheet()->fromArray($result_q3, null, 'A65');

$objPHPExcel->setActiveSheetIndex(0);

// сохранение файла excel

               $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
			   $objWriter->setIncludeCharts(true);			   
               $objWriter->save(dirname(__FILE__)."/dir1"."/".$name);		


			//EMAIL
				//recipient
//				$to = 'puy_t@unitoll.ru;dos-santos_m@unitoll.ru;nikiyan_o@unitoll.ru;erdeli_g@unitoll.ru;krotov_a@unitoll.ru,shvec_k@unitoll.ru;gradoboev_e@unitoll.ru;lemsadze_a@unitoll.ru;poberezhets_m@unitoll.ru;Andrey.Kapitanov@nwcc-msp.ru;borisenkova_n@unitoll.ru;dmitrikova_v@unitoll.ru;fomin_d@unitoll.ru;Aleksey.Chernyshov@msp-highway.ru;mamedov_s@unitoll.ru;taran_e@unitoll.ru';
				$to = 'krotov_a@unitoll.ru,timakov_r@unitoll.ru';
				//sender
				$from = 'do_not_reply@unitoll.ru';
				$fromName = 'TMCO call statistics';

				//email subject
				$subject = 'M11 15-58 TMCO Statistics on the  '.date("d/m/Y",time()-86400); 

				//attachment file path
				$file = (dirname(__FILE__)."/dir1"."/".$name);
				$filewin = ("/mnt"."/smb"."/".$name);
				copy($file,$filewin);

		

				$namero = 'docxfromdb '.date("Y-m-d", time()-86400).'.xlsx';				
				$filero =(dirname(__FILE__)."/rojs"."/".$namero);
				$filewinro = ("/mnt"."/smb"."/".$namero);
				copy($filero,$filewinro);
				

				
				//email body content
				$htmlContent = '<p>Dear All,</p> 
				                <p>Please find in the attachment file with statistics of calls. This message was sent automatically please do not respond to it.</p>
								<p>Best regards</p>';

				//header for sender info
				$headers = "From: $fromName"." <".$from.">";

				//boundary 
				$semi_rand = md5(time()); 
				$mime_boundary = "==Multipart_Boundary_x{$semi_rand}x"; 

				//headers for attachment 
				$headers .= "\nMIME-Version: 1.0\n" . "Content-Type: multipart/mixed;\n" . " boundary=\"{$mime_boundary}\""; 

				//multipart boundary 
				$message = "--{$mime_boundary}\n" . "Content-Type: text/html; charset=\"UTF-8\"\n" .
				"Content-Transfer-Encoding: 7bit\n\n" . $htmlContent . "\n\n"; 

				//preparing attachment
				if(!empty($file) > 0){
					if(is_file($file)){
					$message .= "--{$mime_boundary}\n";
					$fp =    @fopen($file,"rb");
					$data =  @fread($fp,filesize($file));

				@fclose($fp);
				$data = chunk_split(base64_encode($data));
				$message .= "Content-Type: application/octet-stream; name=\"".basename($file)."\"\n" . 
				"Content-Description: ".basename($files[$i])."\n" .
				"Content-Disposition: attachment;\n" . " filename=\"".basename($file)."\"; size=".filesize($file).";\n" . 
				"Content-Transfer-Encoding: base64\n\n" . $data . "\n\n";
				}
				}
				$message .= "--{$mime_boundary}--";
				$returnpath = "-f" . $from;
				//send email
				$mail = @mail($to, $subject, $message, $headers, $returnpath);  
                ?>


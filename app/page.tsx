"use client";

import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";

// ==========================================
// BACKUP DATA 
// ==========================================
const fallbackCategoryMapping: Record<string, string> = {
  "Material tercampur": "Product", "Material salah": "Product", "Material kotor / rusak": "Product",
  "Material terkontaminasi": "Product", "Label hilang / informasi pada label tidak sesuai": "Product",
  "Critical quality parameter issue pada material": "Product", "Critical Visual issue pada material": "Product",
  "Mesin tidak lancar disebabkan kondisi material": "Product", "Dokumen hilang": "Process",
  "Spec salah": "Process", "Tidak ada MO atau MO salah": "Process", "SOP tidak diikuti": "Process",
  "Electronic System bermasalah": "Process", "Kondisi area kerja bermasalah": "Process",
  "Instrument tidak terkalibrasi": "Process", "Online/inline Controls bermasalah": "Process", "Other": "Other"
};

const fallbackOther = [
  "Adhesive Nozzle Area", "Bobin PW", "Body/ Tutup Tray", "Capsule Drum e. Centreline", "Conveyor",
  "Conveyor tray", "Cutting area", "Delivery durm", "Delivery roll", "Flexlink", "FTPU",
  "Garniture area", "Hopper area", "Camera capsule", "Ledger", "Man power", "PZ booth",
  "QTM/ Instrument", "Stuffer jet", "TPP area"
];

// DATA OPERATOR 
const fallbackOperators = [
  "NIRA ASFARINA 000917", "DEKY PRADANA 000940", "ALITTIO F YASSIN 000972", "500833 A. SYAHRONI", "001016 AAN ARIS HIDAYAT", "000322 ABD. ROFIK", "000466 ABDUL CHOLIK KURNIAWAN", "000391 ABDUL MUIS", "000841 ABDURACHMAN A.B", "000637 ABU BAKAR YUSUF", "500967 ABU BAKAR YUSUF", "500839 ACH EFFENDI", "001019 ACHMAD EFENDI", "000949 ACHMAD KHOIRUDIN", "000873 Achmad Ramadhan", "000356 ACHMAD SUHADI", "900001 ACHMAD ZAMRONI", "000695 ADI IRAWAN", "000149 ADI PURNOMO", "500811 ADIN S", "000941 ADIN SULISNO", "500973 ADITYA LESMANA", "000968 ADITYA PRANOTO", "500963 ADRYAN LASTA", "000830 AFIF MAHYUDIN", "000939 AFRITA EKA A", "000634 AFRIYAN BUDHI S", "000991 AGUNG ADI", "500836 AGUNG ADI", "000857 AGUNG EFENDI", "000491 AGUNG PERMADI", "500885 AGUNG SETIAWAN", "900014 Agung Wicaksono Nareswara", "000023 AGUS FAHRONI", "500927 AGUS HARIANTO", "000013 AGUS HARIYANTO", "500001 AGUS MUDJIONO", "500023 AGUS PRAYITNO", "500840 AGUS PURWO AJI", "000328 AGUS SALIM", "500886 AGUS SETYO", "000179 AGUS SULONO", "000462 AGUS ZAINAL FANANI", "000290 AHMAD FASHIHUL LISAN", "500009 AHMAD FUAD", "001005 AHMAD IRFAI", "500938 AHMAD RIYAN FAJAR", "000707 AHMAD WAGIONO", "500887 AHMAD WIDODO", "500007 AHMAD ZAINAL", "00829 AJI PUJONINGRAT", "000986 AKBAR TIGER", "500888 AKBAR TIGER Y.P", "500937 AKHMAD AGUNG", "000661 AKHMAD IRFAI", "900970 AL MUJAHIDIN", "000452 ALEXANDER RADITIO DANANJOYO", "500965 Alfa Rizky Ramadhan", "500812 ALFI DIMO F", "000589 ALFIAN DIVIANSYAH", "001006 ALFIAN IMAM HIZBULLAH", "000419 ALI ARIFIN", "500841 ALI MUKHTAR", "000273 Alif Hidayatullah", "000925 ALIK DINIKAVANILA", "000687 ALIM SAFAROH", "000972 ALITTIO FATAH YASSIN", "500842 AMALIA", "500813 ANA MARDIANA", "000753 ANA ZAKARIA", "500889 ANANDA BAYU", "000777 ANANG JUNAIDI", "000117 ANANG SUBAGYO", "000708 ANANG WAHYUDI", "000938 ANANG WAHYUDI", "000070 ANAWATI", "000397 ANDHIKA HERMANTO", "000298 ANDI JUNIARTO", "000311 ANDI RACHMAN", "000809 ANDI WIBOWO", "500809 ANDI WIBOWO", "000392 ANDIK DWI YULIANTO", "000578 ANDIK PRASETYO", "000650 ANDIK PRIYANTO", "500837 ANDIS SETYAWAN", "000883 ANDRI BAGUS P", "000407 ANDRI SUROSO", "000398 ANGGA BUANA PUTRA", "500815 ANGGA P", "000748 ANGGA PRADITYA", "500968 ANGGA PRASADANA", "000682 ANGGA RISWANDA", "500806 ANGGA RISWANDA", "000647 ANGGARDA ANGGRIAWAN", "000771 ANGGI SAHAT MARTUA", "000145 ANTON HARIYADI", "000312 ANTON KURNIAWAN", "500929 ANUF ARDIANSYAH", "000944 APRIL EFFENDI", "500890 APRIL EFFENDI", "500036 APRILIA", "500891 ARBIAN DEWANTARA", "000995 ARBYAN DEWANTARA", "500915 ARDIANSYAH", "000890 ARDIANSYAH PRASETYO", "500012 ARI", "500892 ARI CAHYONO", "000521 ARI DYAN SETYANINGRUM", "000989 ARI TRI W", "500893 ARI TRI WIJANARKO", "000336 ARIC ASDIADI", "000475 ARIEF ISMAIL", "000449 ARIEF SUDARSONO", "000409 ARIEFIANTO EFENDY", "000575 ARIF ACHMAD ARIS", "001004 ARIF EFENDI", "500894 ARIF EFFENDI", "000089 ARIF HERMAWAN", "000525 ARIF KURNIAWAN", "000571 ARIF LUKMAN", "000697 ARIF MUJI SYUKUR", "000895 ARIF WIBOWO", "500834 ARIF WIBOWO", "000799 ARIFIN", "500799 ARIFIN", "000551 ARIK SUSILO", "500895 ARIS HIDAYATULLAH", "500022 ARIS S", "000632 ARISTYAN DHANI S", "000680 ARIYANTO", "500896 ARIYANTO", "000996 ARIYANTO996", "500969 ARYA GADING KUSUMA", "000715 ASEP BENI", "000590 AYUB DARMAWAN", "000490 BADRUT TAMAM", "500844 BAGUS PRASETYO", "000992 BAGUS RAMADHANI", "500816 BAGUS RAMADHANI", "500845 BAKTIAR BAGUS", "000370 BAMBANG PRABOWO", "500878 Bambang Sulistyono", "000250 BANGKAMID", "001003 BAYU TYO SANDY", "500897 BAYU TYO SANDY", "000237 BBG SULISTYONO", "000465 BENY INDRIAWAN", "000436 BRAMANTYA PRADANA SAPUTRA", "000675 BUCHORI", "000804 BUCHORI", "500804 BUCHORI", "000620 BUDHI PAMUNGKAS", "000683 BUDI ADIANTO", "000891 BUDI IRAWAN", "900002 BUDI PURNOMO", "000676 BUDI SANTOSO", "000712 BUDI SANTOSO", "000532 BUDI WARDOYO", "000472 BUDIONO", "500854 BUDIONO", "000478 CATUR PRASETYO KUNCORO", "000354 CHAVITDONI", "000722 CHOIRUL ANWAR", "000691 CHOIRUL UMAM", "000163 CHOIRUN NISA", "000107 CHUSNIAH HANUM", "000366 DADI SUHANDI", "000271 DADIK SUKAMTO", "001081 Daffa Hidayat", "500817 DANANG ISMAIL", "000677 DANAR BUDI L.", "000380 DARIANTO", "000318 DARYONO", "000964 DEDY OKTA PERMANA", "000660 DEDYK KISWORO", "000940 DEKY PRADANA", "000621 DENI HADI KRISTIAWAN", "000385 DENY ESTIYO", "000518 DENY SETIAWAN", "001021 DEWANGGA ADI TIARNO", "500020 DEWI KARTIKA", "500035 DEWI MA'RUFAH", "500847 DHANI PRASETYO", "000844 DHANU RENDRA", "500916 DIAS P", "500922 DIAS P", "500898 DICTAR", "000573 DIDIK ARIYANTO", "000880 DIDIK IMAM W", "000768 DIDIK KURNIAWAN", "000333 DIDIK PRASETYO", "900008 DIDIK WIDIANTO", "000338 DIDIK YULIANTO", "500028 DIDIK YULIANTO", "500033 DIDIN YUSUF", "500848 DIDIT", "500953 DIDIT AGUSTIYO", "500899 DIMAS ANDHIKA P", "500941 DIMAS BAYU P", "500924 DIMAS MAHENDRA", "500946 DIMAS SURYANA M.", "500005 DINAR R", "500818 DINDA RIZKA M", "000382 DIYA WIRYANTANU", "000017 DJUMILAH", "000278 DODI WIYONO", "000377 DODIK KURNIAWAN S.", "000516 DONY SANJAYA", "000272 DRI SUSANTO", "000825 DWI AGUS YULIANTO", "000441 DWI ARI PURNOMO", "500819 DWI MULYANI", "000359 DWI PRASETYO", "000733 DWI PUJI YUSWANTO", "000488 DWI RAHMAT SUBAGIO", "000868 DWI SUPRIHANTO", "000361 DWIWIT JAYA AGUS PRASETYO", "000550 DYAH PRASASTI POERWA W.", "000471 EDDY WIDIANTO", "500015 EDI GIYANTO", "000450 EDI IWANTORO", "000426 EDI WALOYA", "000901 EDO WILLEM RUMBAY", "500900 EKA DESI", "500801 EKO BUDI P", "000801 EKO BUDI PRASETIAWAN", "000542 EKO CHANDRA WIGUNA", "000303 EKO MARGO SUHARTONO", "000495 EKO SISWANTO", "000247 EKO SUBAKTI", "000383 EKO YULIANTO", "000889 EMIL SALIM", "000692 ENDIK PURGIANTO", "000381 ENDY TRI BAGIO", "500034 ENNY KUSRINI", "900016 ERA ROWATI", "000405 ERAWAN TRI SUDOKO", "500957 ERLANGGA M SAPUTRA", "000820 ESFANING SALIHAT", "500849 EVI SRIANA", "000645 FAKHRUDI", "001064 FAQIH ALFAUZAN", "000988 FARID SUHADAK", "500901 FARID SUHADAK", "000859 FARRIJ ABDILLAH", "500850 FATCHUR ROZZI", "000978 FATKHUL BARRI", "000955 FAUZI OKTAVIANTO", "000504 FAUZI RACHMAN JUNI", "500936 FEBIAN GUSTAV", "000434 FERI TRIWAHYONO", "900022 FERI WARDANI", "000894 FERIANTO", "000507 FERY YULIANTO", "000970 FETI WARDANI", "000737 FIKI FITRIA ARIYANTO", "900019 FITRI SULISTIOWATI", "000390 FREDDY", "000694 FREDDY HASTOMO", "000867 FREDDY HASTOMO", "500944 GALANG PUTRA PERMANA", "500955 GALANG PUTRA PERMANA", "000041 GATOT PRI IRIANTO", "000512 GATOT ROSENO", "000706 GERRY ERLANGGA", "000933 GERRY ERLANGGA", "000345 GHOFUR ABDUL RAHMAN", "000684 GIANTO", "000807 GIANTO", "500807 GIANTO", "500832 GIGIH ANGGA", "000843 GILANG KOPA WIBISONO", "000420 GIYANTO", "000379 GUNTAMA SUKARNO", "500945 GUNTUR FIRMANSYAH", "500954 GUNTUR FIRMANSYAH", "000827 GURUH BRAMANDIKA", "000376 HADI SAPUTRO", "000535 HADI SJAIFULAH", "500902 HADI SUSANTO", "900010 HAFIZ FAJAR S", "500029 HALIMATUS", "001015 HANDOKO", "500919 HANDOKO", "000704 HANDOKO TRI P.", "000501 HANNY", "000310 HANOK WIJATMOKO", "000355 HARDIYANTO", "000389 HARI KUSWANTORO", "000587 HARI NUGROHO", "000618 HARIYANTO ENDRO W", "000138 HARSONO", "000698 HARTADI PURWANTO", "000415 HARTOYO", "001007 HARYA NAGHANTA", "001018 HASAN BISRI", "500903 HASAN BISRI", "000678 HEIRI SUNARKO", "000874 HENDRA ANANTA", "000424 HENDRA KRESA WIJAYA", "000557 HENDRA SAPUTRA", "000375 HENDRI PURNOMO", "000494 HENDRI SETIAWAN SANJAYA", "000567 HENDRO SUNYOTO", "500820 HENGKY SETYO", "500851 HENO P", "500821 HERI P", "000003 HERU PRAPTO SUBAGIYO", "500876 HERU SANTOSO", "000560 HERU SUSANTO", "000534 HERY FIRMANSYAH", "000502 HIBNU ZUBER", "000852 HUSNY ERNANTO", "000515 I WAYAN MARTHA UTAMA", "000374 IAN ANINDYA", "000463 IDRIS SAFII", "000402 IFAN JUNAEDI", "500852 IKLASUL AMALI", "500037 ILHAM DIMAS", "000353 IMAM FACHRUDIN", "500853 IMAM MOCH CHOIRUL", "000559 IMAM SUWIRYO", "000221 IMAM WAHYUDI", "000368 INDARTO DEDDY PRASETYO", "001009 INDIRA ESANTYA AMALIA", "001020 INDRA JUMANTORO", "000509 INDRA LISYANTO", "000075 INSANAH", "000364 IPUNG FEBIANTO", "000723 IRFAN", "500964 Irgi Ahmad Fachrezi", "001017 IRVAN AGUNG", "500904 IRVAN AGUNG S", "500918 ISMAIL JAYANTO", "000699 ISMAIL MARIYANTO", "000705 ISPATMOKO", "500939 ISPATMOKO", "000513 ISRO SUNARTO", "000709 ISTIARTO UTOMO", "500855 ISTIARTO UTOMO", "000998 ITA MELLINIA FEBRIANTI", "000417 IVA YUNIASTUTIK", "000617 IVAN DEKA SURYA", "000946 IVAN DERVANDA", "000717 IWAN HERMANSYAH", "000136 IWAN PRASETYO", "000446 IWAN SETIAWAN", "000530 IWAN SETIAWAN", "000477 IWAN YUDHANTORO SUGIRI", "500008 JAJAK YUDO", "900003 JARROT SULASAH", "000662 JASMARI", "JITMA JITTRANUN MANEEWAN", "000936 JOKO PRAJITNO", "000703 JOKO PRAYITNO", "500880 JOKO PRAYITNO", "000533 JOKO SANTOSO", "000686 JOKO SANTOSO", "000664 JOYO SASMITO", "500974 JUANDA KARUNIA", "000215 JUHAN KUNCAHYONO", "000295 JULI SUSILO", "000529 JUMADI", "500856 JUWARNI", "500920 K. ADI VASELA", "000464 KHASAN ALI", "000024 KHOIRUL HUDA", "500921 KHOIRUL UMAM", "500019 KHUSNUL K", "000734 KOKO KURNIAWAN", "500822 KRISNA SANTI", "000710 KRISTIAWAN", "000846 KUKUH TRIONO SAPUTRO", "000245 KUKUH WAHONO", "000008 KURNIATI", "000963 KURNIAWAN PRASETYO", "000811 KUSBIYANTO", "000470 KUSHARSONO", "000497 KUSNAN", "001033 LELY MAHFIDA", "500823 LIA JANUARSYTA W", "000598 LINDA ERLIA SARI", "000363 LINNY NANIWATI", "000319 LUCKY TRIYANTO", "000547 LUCKY TRYANTO", "000503 LUKMAN HADI NAIK", "000700 LUKMAN HAKIM", "500905 LUKMANUL HAKIM", "900009 LUSIANA", "001062 LUTFI HIDAYAT", "000084 M A Q S U D", "000979 M AMIL MUDZAGI", "500942 M ARIF MUZAKI", "500859 M CHOZIN ASRURI", "500930 M FARIS NASRULLAH", "500860 M FUAD", "500864 M NAAFI", "000856 M NASIR", "500861 M SAFILLAH", "500824 M. ABDUL JALIL", "000999 M. ADITYA KAMANDANU", "000713 M. ALGOZALI", "000967 M. ANDI S", "500825 M. ANDI S", "000638 M. AZIS MUSLIMIN", "500906 M. DIMAS EKA P", "000965 M. EKY KURNIAWAN", "500838 M. EKY KURNIAWAN", "500826 M. FAJAR", "500914 M. FAJAR A.", "500907 M. FATKHUR ROHMAN", "000190 M. FITRUL ANSHORI", "000690 M. KHOIRUL ASHAR", "000679 M. LIYAN SYAHBANA", "500908 M. LYAN", "000977 M. LYAN SYABANA", "000643 M. MAS ZAINURI", "500947 M. NIZAR", "000274 M. NUR AMIN", "000079 M. SAIFUL", "500803 M. SHOKEH", "000672 M. SHOKHEH", "500931 M. WARIZ FANDY", "000987 M. YUWONO", "000701 M. ZAINUL ABIDIN", "000879 M.FURQON", "001000 M.JUNIANTO TRI GUNAWAN", "500928 M.KHAFIDZ HIDAYAT", "000538 M.KHAYYI ZAMRONI", "000984 M.LUKMANUL", "000934 M.ZAINUL", "000625 MAHRUS FAUZI", "001072 MARIA BARINGBING", "000906 MARIYANTI", "500830 MARIYANTI", "000386 MARYONO", "000877 MASNANING YULANDARI", "000951 MASROHAN", "900012 MASRURI HIDAYAT", "000524 MAT SUYANTO", "000427 MEHRU SETIYANTO", "000800 MEI MASHUDI", "500800 MEI MASHUDI", "000648 MIFTACHUL ANWAR", "500862 MIFTAKHUS SUBKHAN", "000969 MOCH BAGUS WICAKSONO", "500966 Moch Edi Kurniawan", "500970 MOCH FAWWAZ JOCELIN", "000326 MOCH. HASYIM", "000558 MOCH. HUTRI", "000945 MOCH. JAMAL TRISWANTO", "500917 MOCH. RIZQY A", "000508 MOCH. ROBAH", "000457 MOCH. ROMLI", "000189 MOCH. SOCHEH", "000444 MOCH.ANIS FAUZI", "000332 MOCH.IMAM FAHRUDIN", "000365 MOCH.JAHJA ISMANTO", "000429 MOCH.KHANDY GAGAR SURYA", "000994 MOCHAMAD RIDWAN", "000199 MOCHAMAD SAIPUL ANWAR", "000305 MOH. SULAIMAN", "000511 MOH. SYAMSUL FUAD", "000948 MOH. ZAINUL ABIDIN", "000746 MOHAMAD SOBIR", "000425 MOHAMMAD JALU PRAYITNO", "000442 MUCHAMAD SOLU", "000962 MUCHLIS IRAWAN", "000151 MUDJI RAHAYU", "500863 MUFIDA", "900011 MUH. FADOLI", "000440 MUHAMAD LUKMAN", "000882 MUHAMMAD HIZBILLAH", "001075 MUHAMMAD IQBAL", "000487 MUJID", "900018 MUJILAH", "000803 MUKHAMMAD SHOKHEH", "000570 MULYADI", "000240 MULYONO", "000990 MULYONO", "500910 MULYONO", "000400 MURIYONO", "000633 MURTAFI RIZQY", "500040 NADA AN'UMILLAH", "500932 NANANG BUDI S.", "000346 NANANG DWI HERNANTO", "000498 NANANG WAHYU WIDODO", "000489 NANGA WAHYUDI", "500018 NANIK S", "000447 NIKOLAUS WINANDAR", "000917 NIRA ASFARINA", "500039 NOVA MAULANA", "000473 NOVI KUSDIANTO", "500031 NOVIA S W", "000644 NOVIE ANAM", "500025 NOVITA", "000430 NUR ARIF SAIFUL", "500030 NUR HABIBAH", "000639 NUR HARIYANTO", "000451 NUR ISMADI WOKO", "000674 NUR MUHADI", "500865 NUR MUHADI", "000517 NUR ROCHIM", "500843 NUR ROHMAD", "000514 NUR YANTO", "500866 NURIANI", "500041 NURITA", "999999 OPERATOR MANDIRI", "500867 PINKY", "000340 PIPIN KRISTIONO", "000016 PONAWAN", "000456 PONGKIE PRASETIA NURCAHYO", "500940 PRAMANA ALIEFIAN AW", "000467 PRASETYO", "000918 PRAVITASARI ANJAR PRATIWI", "000005 PRIAMBODO", "000177 PRISMA UDHIANA", "000406 PULUNG SUHARTONO", "000044 PURNOMO", "500943 PURNOMO DWIANTO", "000616 PURNOMO HARIADI", "000113 PURWANTO", "000439 PUTOT HARIYONO", "500872 PUTRA ERY KW", "001027 Putra Pratama", "500038 PUTRI NILA S", "000546 R.AGUS MARWANTO", "500933 RADIVA THYO S.", "000853 RAHMAN ADETYA", "000667 RAKHMAD HARIYANTO", "000937 RAKHMAD HARIYANTO", "500923 RANGGA ADITIYANSYAH", "000461 RARYK SARPANI TRI HARTANTO", "000433 RATNA SETIAWAN", "500882 RENDRA EKA K", "900007 RENDY HARDRIANA PUTRA", "500873 RHIO HARTANTO", "500956 RIKY ADY DHARMAWAN", "000826 RINGGA HERMAWAN", "500014 RIO HARTANTO", "500032 RISZA MAGHESTI", "500948 RIVAL SYARONI A.", "500925 RIYAN DWI SAPUTRO", "500869 RIZA SURYA", "000862 RIZAL MAHENDRA", "500868 RIZKY DWI", "500934 RIZQI MAULANA F.", "000510 ROBBY IRKHAMNI TANTA", "500002 ROCHIDA", "000663 ROCHMAN YULIANTO", "000931 ROCHMAN YULIANTO", "000537 ROCKIM", "000393 RONI DWI CAHYONO", "000435 RONY ANDIWIANTO", "000898 RORON CATUR", "001001 ROSYITA FIRDAUS", "000408 ROY BAGUS WIDIYANTO", "000327 RUDI ISMANTO", "000718 RUDI ISYANTO", "500835 RUDI WIDAYAT", "000854 RURY NURUL UTAMI", "000437 RYNTO DARMAWAN", "500814 S. ANDRIANTO", "000230 SAIFUDIN", "000666 SAIFUL ANWAR", "000966 SAIFUL BAHRI", "000720 SALIM", "900013 SALUD GAGANG INANSYAH", "000358 SAMSUL ARIFIN", "500827 SANDIKA KUSUMA", "000985 SANDY AGITA RIZKI", "000241 SARDJONO BUDI SANTOSO", "500828 SAROFI", "000549 SELAMET HERU KARNOTO", "000459 SENO ADJIE", "000552 SEPDIANTO SUKO PRASETYO", "500949 SEPTIAN ARI P.", "000384 SETIA HIDAYAT", "500021 SETIABUDI", "000522 SETYO ANDI KURNIAWAN", "500971 SHAFA SHEVANIA ARIP", "500004 SHANDY YOGA", "000688 SIGIT PRAJITNO", "000932 SIGIT PRAJITNO", "500871 SILVI", "000067 SISWOKO", "000330 SISWOKO", "500911 SITI MUALIFATUL FITRI", "000822 SITI ROMLAH NURUL BADRIYAH", "000952 SITI ZAZIMAH", "000569 SLAMET HARYADI", "000711 SLAMET RIYADI", "000523 SLAMET WIDODO", "000002 SOEGENG SOEPRIYANTO", "000014 SOEGIONO", "500912 SRI SUNDARI", "900015 SRIWATI", "000849 SRIYADI", "900017 SRIYANI", "000454 SUBAGYO", "500950 SUBAKTIAR JUNI P.", "000156 SUBINI", "000195 SUDARJONO", "500016 SUDARNO", "000371 SUDARWANTO", "000200 SUDARWOKO", "000022 SUDIRMAN WIDIARTO", "000476 SUGENG HANDRIANTO", "000428 SUGENG TRIYONO", "000105 SUGIHARTOYO", "000545 SUHARTONO", "000572 SUJATMIKO", "000577 SUJATMIKO", "000040 SUKADI", "000689 SUKAMTO", "000808 SUKAMTO", "500808 SUKAMTO", "000659 SUKARNI", "000343 SUKARNO", "000527 SUKI", "000492 SUKIRAN", "000805 SUKOCO", "500805 SUKOCO", "000484 SUKURIYO NUGROHO", "000526 SULARNO", "500913 SULIS SUBANDIO", "000347 SULISTIYONO", "000080 SULISTYO WIDODO", "000531 SULIYONO", "000029 SUMARDI", "000413 SUMARIYONO", "000588 SUMINTO", "900020 SUMIRAH", "000866 SUNARTO", "900004 SUPARDI", "000443 SUPARI", "000157 SUPARMAN ARIEF", "000037 SUPARMI", "000082 SUPARTI", "000257 SUPRIONO", "000702 SUTIKNO", "000993 SUTIKNO", "000496 SUTIMIN", "000520 SUTOTO HEXAMBONO", "000421 SUTRISNO", "000553 SUWARI", "000576 SUYANTO", "000258 SUYONO", "000714 SUYONO", "000685 SUYUD HUDAYAH", "000404 SYAHFIRI SYAMSYASA", "000556 SYAIFUDIN", "000468 SYAIFUL MUHYIDIN", "1008 SYAIFUL MUHYIDIN", "500972 SYLFIA", "99999 System", "000184 TAJUDIN AHMAD", "500874 TATON", "500951 TAUFIK ELSA H", "000244 TEGUH MARUTO", "000460 TEUKU FAISAL UMRY", "000030 TINGKAT", "000438 TONY YOEDIANTO", "000108 TOTOK SUSANTO", "000344 TOTOK YONI PURWANTO", "000479 TOTOK YUWONO", "000528 TRI AGUNG PAMBUDI", "500875 TRI AGUS P", "000297 TRI JATMIKO", "000057 TRI RUDIYANTO", "000474 TRI WAHYUDI PRASETYO", "000369 TRIAJI BUDI PRIMAWAN", "000860 TRIAJI BUDI PRIMAWAN", "000351 TRY WAHYUDI", "500829 TURMIYAH", "000388 UMAR FADLY", "500026 UMASHANKAR", "000499 VICTOR DEDY INDRAJAYA", "000396 VIDO AGUS PRASETYO", "000544 VIKI ARIANTO", "500027 VIRUPAKSHI", "500935 WAHONO", "000767 WAHYU ARDIYANTO", "900005 WAHYU NUGROHO", "000884 WAHYU PURNOMO", "000892 WAHYU PURNOMO", "000899 WAHYU PURNOMO", "000485 WAHYUDI SULISTIO AMERTANI", "000665 WAKIT MOHARI", "000416 WARSINO", "000350 WAWAN SISTIYONO", "000455 WAYAN FIRLY SETIAWAN", "000548 WIBOWO APRIANTO", "000693 WIDE MAHENDRA", "000865 WIDE MAHENDRA", "000090 WIDUK WAHYUBI", "000997 WIMA RAMADHAN", "000414 WIMA RAMADHANA", "000902 WURYANTO", "000561 YANUAR ADI RUPAWAN", "000574 YANUAR PRAMIANTO", "000719 YARIANTO", "900021 YEKTI LESTARI", "000885 YENDI RACHMADY", "000540 YHUDI YHULISTIKO", "500011 YOGA P", "000961 YOGI MADA", "500810 YOGI MADA", "000453 YOHAN INDRAWIJAYA", "000378 YOHAN NOERDIANTOKO", "000716 YONO ARI", "000352 YOYOK HADI PURWANTO", "000445 YUARDINATA", "000971 YULI RAHMAWATI", "900023 YULI RAHMAWATI", "500024 YULIANA", "500003 YULIATIN", "500831 YULIATIN", "000337 YUNAIDI SETIAWAN", "000506 YUNAIDI SETIAWAN", "500879 YURIKE ELOK PURWANTI", "000394 YUSUF AHMAD", "000696 YUSUF TRIANTO", "000403 ZAENAL AFANDI", "000162 ZAHRO", "000367 ZAINAL ABIDIN", "500883 ZAINUL ABIDIN", "000036 ZAKARIA", "500952 ZAKARIA ALFIANDI", "500962 ZAKARIA ALFIANDI"
];

export default function QualityAlertDashboard() {
  // --- NAVIGATION STATE ---
  const [activeView, setActiveView] = useState("Data Input");

  // --- STATE DATA MASTER ---
  const [alertCategoryMapping, setAlertCategoryMapping] = useState<Record<string, string>>(fallbackCategoryMapping);
  const [otherOptions, setOtherOptions] = useState<string[]>(fallbackOther);
  const [operatorOptions, setOperatorOptions] = useState<string[]>(fallbackOperators);
  const [excelStatus, setExcelStatus] = useState("🔄 Menyiapkan data...");

  // --- STATE MANAGEMENT (FORM) ---
  const [alertTypeBtn, setAlertTypeBtn] = useState("Product");
  const [shiftBtn, setShiftBtn] = useState("1");
  const [typeAlertSelect, setTypeAlertSelect] = useState("");
  const [inputMesin, setInputMesin] = useState("");
  const [teamInfo, setTeamInfo] = useState({ name: "Terverifikasi otomatis", bg: "#ffffff", color: "#64748b", border: "#cbd5e1" });
  
  // Custom Autocomplete State (MEMPERCEPAT PENCARIAN & FITUR "OTHER")
  const [initiator, setInitiator] = useState("");
  const [actionOp, setActionOp] = useState("");
  const [otherSelect, setOtherSelect] = useState(""); // <-- State baru untuk Other
  const [showInitiatorDropdown, setShowInitiatorDropdown] = useState(false);
  const [showActionOpDropdown, setShowActionOpDropdown] = useState(false);
  const [showOtherDropdown, setShowOtherDropdown] = useState(false); // <-- Dropdown Other

  // Form Inputs
  const [tglTemuan, setTglTemuan] = useState("");
  const [wktTemuan, setWktTemuan] = useState("");
  const [tglCorr, setTglCorr] = useState("");
  const [wktCorr, setWktCorr] = useState("");
  const [moNum, setMoNum] = useState("");
  const [itemCode, setItemCode] = useState("");
  const [eventDesc, setEventDesc] = useState("");
  const [corrAction, setCorrAction] = useState("");

  const [mainData, setMainData] = useState<any[]>([]);

  // ====================================================================
  // LOGIKA DARI FILE EXCEL
  // ====================================================================
  useEffect(() => {
    const fetchMasterDataFromExcel = async () => {
      try {
        const response = await fetch('/Data.xlsx');
        if (!response.ok) {
          setExcelStatus("🟡 Data Siap (Memakai Backup Lokal)");
          return;
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // 1. SAMBUNGKAN KE SHEET "Kategori"
        if (workbook.SheetNames.includes("Kategori")) {
          const sheetKategori = workbook.Sheets["Kategori"];
          const dataKategori = XLSX.utils.sheet_to_json(sheetKategori, { header: 1 }) as any[][];
          
          const mappingObj: Record<string, string> = {};
          const othersArr: string[] = [];

          for (let i = 1; i < dataKategori.length; i++) {
            const row = dataKategori[i];
            if (!row) continue;
            const kolomA_alert = row[0]; // Kolom A
            const kolomB_type = row[1];  // Kolom B
            const kolomC_other = row[2]; // Kolom C
            
            if (kolomA_alert && kolomB_type) {
              let typeVal = "Other";
              if (String(kolomB_type).toLowerCase().includes("product")) typeVal = "Product";
              else if (String(kolomB_type).toLowerCase().includes("process")) typeVal = "Process";
              mappingObj[String(kolomA_alert).trim()] = typeVal;
            }

            if (kolomC_other) {
              othersArr.push(String(kolomC_other).trim());
            }
          }
          if (Object.keys(mappingObj).length > 0) setAlertCategoryMapping(mappingObj);
          if (othersArr.length > 0) {
            setOtherOptions(othersArr); 
          }
        }

        // 2. Operator
        if (workbook.SheetNames.includes("Operator")) {
          const sheetOperator = workbook.Sheets["Operator"];
          const dataOperator = XLSX.utils.sheet_to_json(sheetOperator, { header: 1 }) as any[][];
          
          const operatorsArr: string[] = [];
          for (let i = 1; i < dataOperator.length; i++) {
            const row = dataOperator[i];
            if (row && row[0]) { operatorsArr.push(String(row[0]).trim()); }
          }
          if (operatorsArr.length > 50) {
              setOperatorOptions(operatorsArr);
          }
        }
        setExcelStatus("🟢 Terhubung Otomatis ke Data.xlsx!");

      } catch (error) {
        setExcelStatus("🔴 Gagal konek Excel. Memakai backup.");
      }
    };

    fetchMasterDataFromExcel();
  }, []);

  // AUTO-READ DATABASE TRANSAKSI LEWAT API
  useEffect(() => {
    fetch('/api/data')
      .then(res => res.json())
      .then(data => {
        if (Array.isArray(data)) {
           const cleanData = data.map(d => {
              let dateStr = d.tglT;
              if (typeof dateStr === 'number') {
                  const date = new Date(Math.round((dateStr - 25569) * 86400 * 1000));
                  dateStr = date.toISOString().split('T')[0];
              }
              let timeStr = d.wktT;
              if (typeof timeStr === 'number') {
                  if (timeStr < 1) { 
                      const totalSeconds = Math.round(timeStr * 86400);
                      const h = Math.floor(totalSeconds / 3600).toString().padStart(2, '0');
                      const m = Math.floor((totalSeconds % 3600) / 60).toString().padStart(2, '0');
                      timeStr = `${h}:${m}`;
                  } else { 
                      timeStr = timeStr.toString().padStart(2, '0') + ':00';
                  }
              } else if (typeof timeStr === 'string' && timeStr !== '-' && !timeStr.includes(':')) {
                  timeStr = timeStr.replace('.', ':');
                  if(!timeStr.includes(':')) timeStr = timeStr.padStart(2, '0') + ':00';
              }
              return { ...d, tglT: dateStr || '-', wktT: timeStr || '-' }
           });
           
           const validData = cleanData.filter(d => d.tglT !== '-' || d.alert !== '-');
           const sortedData = validData.sort((a, b) => {
              // Pakai perbandingan string tanggal
              if (a.tglT > b.tglT) return -1;
              if (a.tglT < b.tglT) return 1;
              return 0;
           });
           setMainData(sortedData); 
        }
      })
      .catch(err => console.error("Error database:", err));
  }, []);

  const handleDelete = async (idToDelete: number) => {
    const isConfirm = window.confirm("Apakah Anda yakin ingin menghapus log ini?");
    if (!isConfirm) return;
    const updatedData = mainData.filter(item => item.id !== idToDelete);
    setMainData(updatedData);
    try { await fetch(`/api/data?id=${idToDelete}`, { method: 'DELETE' }); } catch (error) {}
  };
  
  const [dateRange, setDateRange] = useState<{start: string | null, end: string | null}>({ start: null, end: null });
  const [calCurrentDate, setCalCurrentDate] = useState(new Date()); 
  const [showMonthPicker, setShowMonthPicker] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  // LOGIKA FILTER TANGGAL STRING COMPARISON (Lebih akurat & gak kena bug Timezone)
  const filteredData = mainData.filter(d => {
      if (!dateRange.start) return true;
      if (!dateRange.end) return d.tglT === dateRange.start;
      return d.tglT >= dateRange.start && d.tglT <= dateRange.end;
  });

  const handleExportExcel = () => {
    if(filteredData.length === 0) return alert("Tidak ada data untuk diexport!");
    const excelData = filteredData.map(row => ({
      "Alert Type": row.alertType, "Shift": row.shift, "Discovery Date": row.tglT, "Discovery Time": row.wktT,
      "Action Date": row.tglC !== "-" ? row.tglC : "", "Action Time": row.wktC !== "-" ? row.wktC : "",
      "MO": row.mo, "Item Code": row.item, "Machine": row.msn, "Team": row.tm,
      "Initiator": row.opr, "Alert": row.alert, "Event Description": row.desc,
      "Action Description": row.act, "Action by": row.actOp
    }));
    const rangeLabel = dateRange.start ? `${dateRange.start}${dateRange.end ? '_sampai_'+dateRange.end : ''}` : "All";
    const ws = XLSX.utils.json_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Database Logs");
    XLSX.writeFile(wb, `Quality_Alert_Logs_${rangeLabel}.xlsx`);
  };

  const masterDataMesin = {
    "MONO": ["MC", "IM01", "IM02", "IM03", "KM04", "KM10", "KM53", "KM54", "KM55", "KM57", "KR02", "KR01", "KM65", "KM64", "KM67"],
    "CAPSULE A": ["MC", "AS07", "KS01", "KS02", "KS04", "KS06", "KS09", "SL02", "KM62", "TH08", "KM08", "JD02", "FDU1", "FDU2", "FDU3"],
    "CAPSULE B": ["MC", "KS03", "KS07", "KS08", "TH09", "DD10", "DD51", "KM56", "KM58", "KM03", "KM51", "KC05", "KS05"],
    "ND3H": ["MC", "KC07", "KC08", "KC54", "KM63", "KN51", "KN53", "ND01", "ND02", "ND04", "ND05", "ND06", "ND08", "ND07", "SM04", "SM08"],
    "D3H": ["MC", "KC03", "KM13", "KM15", "KM59", "SM01", "SM05", "DD15", "DD07", "VD04", "VD05", "TH10", "DD12", "PT01"]
  };

  const handleMesinChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const value = e.target.value.toUpperCase().trim();
    setInputMesin(value);
    let foundTeam = "Terverifikasi otomatis";
    let bg = "#ffffff", color = "#64748b", border = "#cbd5e1";
    for (const [team, mesinList] of Object.entries(masterDataMesin)) {
      if (mesinList.includes(value)) { foundTeam = team; bg = "#dcfce7"; color = "#166534"; border = "#86efac"; break; }
    }
    setTeamInfo({ name: foundTeam, bg, color, border });
  };

  const handleTypeAlertChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const val = e.target.value;
    setTypeAlertSelect(val);
    if (val && alertCategoryMapping[val] && alertCategoryMapping[val] !== "Other") {
      setAlertTypeBtn(alertCategoryMapping[val]);
    }
  };

  const submitData = async () => {
    if (!tglTemuan || !inputMesin || teamInfo.name === "Terverifikasi otomatis") { 
      alert("Tanggal Temuan dan Kode Mesin valid wajib diisi!"); return; 
    }
    let finalAlertType = typeAlertSelect;
    if (!finalAlertType || finalAlertType.includes("Drop down")) finalAlertType = "-";
    
    // Kalau pilih Other, masukin teks dari input search Other-nya
    if (finalAlertType === "Other") finalAlertType = otherSelect || "Other"; 

    const newData = {
      id: mainData.length > 0 ? Math.max(...mainData.map(d=>d.id)) + 1 : 1,
      alertType: alertTypeBtn, shift: shiftBtn,
      tglT: tglTemuan, wktT: wktTemuan || "-", tglC: tglCorr || "-", wktC: wktCorr || "-",
      mo: moNum || "-", item: itemCode || "-",
      msn: inputMesin, tm: teamInfo.name, opr: initiator || "-", 
      alert: finalAlertType !== "-" ? finalAlertType : "-",
      desc: eventDesc || "-", act: corrAction || "-", actOp: actionOp || "-"
    };

    const updatedData = [newData, ...mainData].sort((a, b) => {
      if (a.tglT > b.tglT) return -1;
      if (a.tglT < b.tglT) return 1;
      return 0;
    });
    
    setMainData(updatedData); // LANGSUNG UPDATE STATE!
    
    try { await fetch('/api/data', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(newData) }); } catch (error) {}

    // Reset Form
    setInputMesin(""); setTeamInfo({ name: "Terverifikasi otomatis", bg: "#ffffff", color: "#64748b", border: "#cbd5e1" });
    setTypeAlertSelect(""); setTglTemuan(""); setWktTemuan(""); setTglCorr(""); setWktCorr("");
    setMoNum(""); setItemCode(""); setEventDesc(""); setCorrAction(""); setInitiator(""); setActionOp(""); setOtherSelect("");
    setCurrentPage(1); 
  };

  const totalPages = Math.max(1, Math.ceil(filteredData.length / itemsPerPage));
  const startIdx = (currentPage - 1) * itemsPerPage;
  const pageData = filteredData.slice(startIdx, startIdx + itemsPerPage);

  const masterDefects = {
      CAPSULE: [ "Other", "Garniture Area", "FTPU", "Kondisi area kerja bermasalah", "Material kotor / rusak", "Adhesive Nozzle area", "Capsule Drum", "Critical Visual issue pada material", "PZ Booth", "TPP Area", "Electronic System bermasalah", "Cutting Area", "Bobin PW", "Delivery Roll", "Delivery Drum", "Critical quality parameter issue pada material", "Label hilang / informasi pada label tidak sesuai", "Stuffer Jet", "Centreline", "QTM / Instrument", "Ledger", "Mesin tidak lancar disebabkan kondisi material", "Conveyor", "Material terkontaminasi", "Tidak ada MO atau MO salah", "Body / Tutup Tray", "Kamera Capsule" ],
      MONO: [ "Body / Tutup Tray", "Other", "Material kotor / rusak", "TPP Area", "Centreline", "Delivery Drum", "Critical Visual issue pada material", "Garniture Area", "Electronic System bermasalah", "Cutting Area", "QTM / Instrument", "Adhesive Nozzle area", "FTPU", "Conveyor" ],
      ND3H: [ "Hopper Area", "Critical Visual issue pada material", "Garniture Area", "Critical quality parameter issue pada material", "Bobin PW", "Adhesive Nozzle area", "Other", "Cutting Area", "Conveyor", "Delivery Drum", "Mesin tidak lancar disebabkan kondisi material", "QTM / Instrument", "Conveyor tray", "Kondisi area kerja bermasalah", "Tidak ada MO atau MO salah", "FTPU" ]
  };

  const normalizeStr = (str: string) => str.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  const generateSummary = (teamKeyword: string, defectList: string[]) => {
    let teamData = filteredData.filter(d => d.tm && d.tm.toUpperCase().includes(teamKeyword.toUpperCase()));
    let counts: { [key: string]: number } = {};
    defectList.forEach(d => counts[d] = 0); 
    let total = 0;
    teamData.forEach(d => {
      let alertValue = (d.alert && d.alert !== "-") ? d.alert : (d.desc || "Other");
      let matchedKey = "Other";
      for(let key of defectList) {
          if(normalizeStr(alertValue) === normalizeStr(key) || normalizeStr(alertValue).includes(normalizeStr(key)) || normalizeStr(key).includes(normalizeStr(alertValue))) {
              matchedKey = key; break;
          }
      }
      counts[matchedKey] = (counts[matchedKey] || 0) + 1;
      total++;
    });
    let summaryArr = Object.keys(counts).map(key => ({ label: key, count: counts[key] })).sort((a,b) => b.count - a.count); 
    let cumSum = 0;
    return summaryArr.map(item => {
      cumSum += item.count;
      let pctBulat = total > 0 ? Math.round((item.count / total) * 100) : 0;
      let cumPctBulat = total > 0 ? Math.round((cumSum / total) * 100) : 0;
      return { ...item, percent: pctBulat + '%', percentNum: pctBulat, cumPercent: cumPctBulat };
    });
  };

  const monoSummary = useMemo(() => generateSummary("MONO", masterDefects.MONO), [filteredData]);
  const capsuleSummary = useMemo(() => generateSummary("CAPSULE", masterDefects.CAPSULE), [filteredData]);
  const nd3hSummary = useMemo(() => generateSummary("D3H", masterDefects.ND3H), [filteredData]); 

  const handleDateClick = (fullDate: string) => {
      if (!dateRange.start || (dateRange.start && dateRange.end)) { setDateRange({ start: fullDate, end: null });
      } else {
          if (fullDate < dateRange.start) { setDateRange({ start: fullDate, end: dateRange.start }); } 
          else { setDateRange({ start: dateRange.start, end: fullDate }); }
      }
      setCurrentPage(1);
  };

  const isDateSelected = (dateStr: string) => {
      if (!dateRange.start) return false;
      if (!dateRange.end) return dateStr === dateRange.start;
      return dateStr >= dateRange.start && dateStr <= dateRange.end;
  };

  const renderCalendarDays = () => {
    const year = calCurrentDate.getFullYear();
    const month = calCurrentDate.getMonth();
    const mStr = (month + 1).toString().padStart(2, '0');
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    let cells = [];
    for (let i = 0; i < firstDay; i++) { cells.push(<div key={`empty-start-${i}`} style={{ height: '32px' }}></div>); }
    for (let i = 1; i <= daysInMonth; i++) {
      let d = i.toString().padStart(2, '0');
      let fullDate = `${year}-${mStr}-${d}`; 
      let isActive = isDateSelected(fullDate);
      let isEndpoint = fullDate === dateRange.start || fullDate === dateRange.end;

      cells.push(
        <div key={fullDate} onClick={() => handleDateClick(fullDate)} style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '32px', cursor: 'pointer', borderRadius: '8px', backgroundColor: isActive ? (isEndpoint ? '#f97316' : '#ffedd5') : 'transparent', color: isActive ? (isEndpoint ? 'white' : '#ea580c') : '#334155', fontWeight: isActive ? 'bold' : '500', transition: 'all 0.2s', border: isEndpoint ? 'none' : '1px solid transparent' }} onMouseOver={(e) => { if(!isActive) { e.currentTarget.style.backgroundColor = '#ffffff'; e.currentTarget.style.borderColor = '#cbd5e1'; } }} onMouseOut={(e) => { if(!isActive) { e.currentTarget.style.backgroundColor = 'transparent'; e.currentTarget.style.borderColor = 'transparent'; } }}>
          {i}
        </div>
      );
    }
    const remainingCells = 42 - (firstDay + daysInMonth);
    for (let i = 0; i < remainingCells; i++) { cells.push(<div key={`empty-end-${i}`} style={{ height: '32px' }}></div>); }
    return cells;
  };

  const ParetoChart = ({ title, data }: { title: string, data: any[] }) => {
    if (data.length === 0) return ( <div style={{ flex: 1, minWidth: '350px', background: 'white', padding: '15px', borderRadius: '16px', border: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', height: '350px', boxShadow: '0 4px 12px rgba(0,0,0,0.05)' }}> <p style={{ color: '#94a3b8', fontWeight: '600' }}>Belum ada data masalah untuk {title}.</p> </div> );
    
    const topData = data.filter(d => d.count > 0 || d.label === "Other").slice(0, 15); 
    const maxCount = Math.max(...topData.map(d => d.count)) * 1.15 || 1; 
    const chartHeight = 200; 
    const barWidth = 100 / topData.length;

    return (
      <div style={{ flex: 1, minWidth: '350px', background: 'white', padding: '25px 30px', borderRadius: '16px', boxShadow: '0 4px 12px rgba(0,0,0,0.05)', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column' }}>
        <h3 style={{ textAlign: 'center', margin: '0 0 25px 0', fontSize: '16px', color: '#1e3a8a', fontWeight: '900', fontFamily: 'system-ui, -apple-system, sans-serif', textTransform: 'uppercase', letterSpacing: '0.5px' }}> Pareto Defect - {title} </h3>
        
        <div style={{ position: 'relative', height: `${chartHeight}px`, width: '100%', borderBottom: '2px solid #94a3b8', borderLeft: '2px solid #cbd5e1', borderRight: '2px solid #cbd5e1', display: 'flex', marginTop: '10px' }}>
          
          <div style={{ position: 'absolute', left: '-35px', top: '-6px', bottom: 0, display: 'flex', flexDirection: 'column', justifyContent: 'space-between', fontSize: '10px', fontWeight: '700', color: '#64748b', alignItems: 'flex-end', paddingRight: '6px', width: '35px' }}>
            <span>{Math.round(maxCount)}</span><span>{Math.round(maxCount * 0.75)}</span><span>{Math.round(maxCount * 0.5)}</span><span>{Math.round(maxCount * 0.25)}</span><span>0</span>
          </div>

          <svg style={{ width: '100%', height: '100%', overflow: 'visible' }}>
            <defs>
              <linearGradient id={`barGradient-${title.replace(/\s/g,'')}`} x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor="#3b82f6" /><stop offset="100%" stopColor="#1e3a8a" />
              </linearGradient>
            </defs>

            {[0.25, 0.5, 0.75, 1].map(pct => (
              <line key={pct} x1="0" y1={chartHeight * (1-pct)} x2="100%" y2={chartHeight * (1-pct)} stroke="#f1f5f9" strokeWidth="1.5" strokeDasharray="5 5" />
            ))}

            {topData.map((item, i) => {
              const bHeight = (item.count / maxCount) * chartHeight;
              return (
                <g key={`bar-${i}`}>
                  <rect x={`${(i * barWidth) + (barWidth*0.1)}%`} y={chartHeight - bHeight} width={`${barWidth*0.8}%`} height={bHeight} fill={`url(#barGradient-${title.replace(/\s/g,'')})`} rx="3" style={{ transition: 'all 0.6s cubic-bezier(0.4, 0, 0.2, 1)' }} />
                  {item.count > 0 && ( <text x={`${(i * barWidth) + (barWidth/2)}%`} y={chartHeight - bHeight - 6} fill="#0f172a" fontSize="11" fontWeight="800" textAnchor="middle" style={{ transition: 'all 0.6s' }}>{item.count}</text> )}
                </g>
              );
            })}
            
            <polyline points={topData.map((item, i) => `${(i * barWidth) + (barWidth/2)}%,${chartHeight - (item.cumPercent/100 * chartHeight)}`).join(' ')} fill="none" stroke="#ea580c" strokeWidth="2.5" style={{ transition: 'all 0.6s cubic-bezier(0.4, 0, 0.2, 1)' }} />
            {topData.map((item, i) => (
              <circle key={`dot-${i}`} cx={`${(i * barWidth) + (barWidth/2)}%`} cy={chartHeight - (item.cumPercent/100 * chartHeight)} r="4.5" fill="#ea580c" stroke="#ffffff" strokeWidth="2" style={{ transition: 'all 0.6s cubic-bezier(0.4, 0, 0.2, 1)' }} />
            ))}
          </svg>

          <div style={{ position: 'absolute', right: '-40px', top: '-6px', bottom: 0, display: 'flex', flexDirection: 'column', justifyContent: 'space-between', fontSize: '10px', fontWeight: '800', color: '#ea580c', alignItems: 'flex-start', paddingLeft: '6px', width: '40px' }}>
            <span>100%</span><span>75%</span><span>50%</span><span>25%</span><span>0%</span>
          </div>
        </div>

        <div style={{ position: 'relative', height: '110px', width: '100%', marginTop: '15px' }}>
          {topData.map((item, i) => (
            <div key={`label-${i}`} style={{ position: 'absolute', left: `${(i * barWidth) + (barWidth/2)}%`, top: 0, width: '0' }}>
              <div style={{ transform: 'rotate(-45deg)', transformOrigin: 'top right', width: '120px', textAlign: 'right', paddingRight: '8px', fontSize: '10px', color: '#334155', fontWeight: '600', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', position: 'absolute', right: 0 }}>
                {item.label}
              </div>
            </div>
          ))}
        </div>

        <div style={{ display: 'flex', width: '100%', borderTop: '2px solid #f1f5f9', paddingTop: '10px' }}>
          {topData.map((item, i) => (
            <div key={`pct-${i}`} style={{ width: `${barWidth}%`, textAlign: 'center', fontSize: '11px', fontWeight: '900', color: '#ea580c' }}>{item.percent}</div>
          ))}
        </div>
      </div>
    );
  };

  // Logika memfilter hasil maksimal (untuk dropdown)
  const filteredInitiators = operatorOptions.filter(op => op.toLowerCase().includes(initiator.toLowerCase()));
  const filteredActionOps = operatorOptions.filter(op => op.toLowerCase().includes(actionOp.toLowerCase()));
  const filteredOther = otherOptions.filter(opt => opt.toLowerCase().includes(otherSelect.toLowerCase())); // <-- Filter Other

  return (
    <div className="main-container" style={{ width: '100vw', height: '100vh', display: 'flex', flexDirection: 'column', fontFamily: "'Poppins', sans-serif", backgroundColor: '#f4f7fb', backgroundImage: "linear-gradient(rgba(244, 247, 251, 0.88), rgba(244, 247, 251, 0.88)), url('/map.svg')", backgroundSize: 'cover', backgroundPosition: 'center', backgroundRepeat: 'no-repeat', overflow: 'hidden' }}>
      
      <style dangerouslySetInnerHTML={{__html: `
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800;900&display=swap');
        body, html { margin: 0; padding: 0; }
        .custom-scroll::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scroll::-webkit-scrollbar-track { background: transparent; }
        .custom-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
        .custom-scroll::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        .input-ui { font-size: 13px; padding: 10px 12px; border-radius: 8px; border: 1px solid #cbd5e1; width: 100%; box-sizing: border-box; outline: none; transition: all 0.2s; color: #1e293b; background: white; font-family: 'Poppins', sans-serif; pointer-events: auto; }
        .input-ui:focus { border-color: #f97316; box-shadow: 0 0 0 3px rgba(249, 115, 22, 0.15); }
        .input-ui::placeholder { color: #94a3b8; }
        .card-row { background: white; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border: 1px solid #f1f5f9; transition: all 0.2s; }
        .card-row:hover { transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.08); border-color: #e2e8f0; }
        .btn-menu { padding: 14px; border-radius: 10px; cursor: pointer; font-weight: 800; font-size: 14px; text-align: center; transition: all 0.2s; border: 1px solid transparent; font-family: 'Poppins'; pointer-events: auto; }
        .btn-menu:hover { transform: scale(1.02); border-color: #cbd5e1; }
        .btn-simpan { background: linear-gradient(135deg, #f97316 0%, #ea580c 100%); transition: all 0.2s; font-family: 'Poppins', sans-serif; color: white; pointer-events: auto; }
        .btn-simpan:hover { transform: translateY(-2px); box-shadow: 0 6px 12px rgba(249, 115, 22, 0.3); }
        .btn-export { padding: 8px 14px; border-radius: 8px; font-weight: 700; font-size: 13px; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; gap: 6px; border: 1px solid #cbd5e1; background: white; color: #334155; font-family: 'Poppins', sans-serif; pointer-events: auto; }
        .btn-export:hover { background: #f8fafc; border-color: #94a3b8; }
        .page-transition { animation: fadeInScale 0.4s ease-out forwards; pointer-events: auto; }
        @keyframes fadeInScale { 0% { opacity: 0; transform: translateY(15px); } 100% { opacity: 1; transform: translateY(0); } }
        .dashboard-layout { display: flex; flex: 1; padding: 20px; gap: 20px; overflow-x: hidden; overflow-y: hidden; width: 100%; box-sizing: border-box; }
        @media (max-width: 1024px) {
            .dashboard-layout { flex-direction: column; overflow-y: auto; overflow-x: hidden; }
            .sidebar-area { width: 100% !important; flex-direction: row !important; flex-wrap: wrap; }
            .sidebar-area > div { flex: 1; min-width: 250px; }
            .input-section { width: 100% !important; max-width: 100% !important; }
            .print-area-container { min-width: 100% !important; overflow-x: auto; }
            .card-row { flex-direction: column; gap: 10px; }
            .card-row > div { width: 100% !important; padding-right: 0 !important; }
        }
      `}} />

      {/* ================= 1. HEADER ================= */}
      <header style={{ 
        position: 'relative', width: '100%', background: 'rgba(255, 255, 255, 0.95)', backdropFilter: 'blur(10px)',
        borderBottom: '5px solid #f97316', boxShadow: '0 4px 15px rgba(0, 0, 0, 0.08)', display: 'flex', justifyContent: 'space-between', 
        alignItems: 'center', padding: '15px 40px', zIndex: 9999, flexShrink: 0, boxSizing: 'border-box'
      }}>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <img src="/logo.png" alt="Logo Filtrona" style={{ height: '45px', objectFit: 'contain' }} onError={(e) => e.currentTarget.style.display = 'none'} />
          <div style={{ fontSize: '26px', fontWeight: '900', color: '#1e3a8a', marginLeft: '15px', letterSpacing: '1px', fontFamily: 'system-ui, -apple-system, sans-serif' }}>
            | DASHBOARD QUALITY ALERT
          </div>
        </div>
      </header>

      {/* LAYOUT CONTAINER: Force Z-Index 9999 + Pointer Events Auto */}
      <div className="dashboard-layout" style={{ position: 'relative', zIndex: 9999, pointerEvents: 'auto' }}>
        
        {/* ================= 2. SIDEBAR ================= */}
        <aside className="sidebar-area custom-scroll" style={{ width: '280px', display: 'flex', flexDirection: 'column', gap: '15px', flexShrink: 0, position: 'relative', zIndex: 9999, pointerEvents: 'auto' }} >
          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', background: '#eef2f6', padding: '15px', borderRadius: '16px', border: '1px solid #cbd5e1' }}>
            <button className="btn-menu" onClick={() => setActiveView("Data Input")} style={{ background: activeView === "Data Input" ? '#1e3a8a' : 'white', color: activeView === "Data Input" ? 'white' : '#475569', boxShadow: activeView === "Data Input" ? '0 4px 8px rgba(30, 58, 138, 0.25)' : '0 1px 3px rgba(0,0,0,0.05)' }}>DATA INPUT</button>
            <button className="btn-menu" onClick={() => setActiveView("Data Management")} style={{ background: activeView === "Data Management" ? '#1e3a8a' : 'white', color: activeView === "Data Management" ? 'white' : '#475569', boxShadow: activeView === "Data Management" ? '0 4px 8px rgba(30, 58, 138, 0.25)' : '0 1px 3px rgba(0,0,0,0.05)' }}>DATA MANAGEMENT</button>
            <button className="btn-menu" onClick={() => setActiveView("Data Analytics")} style={{ background: activeView === "Data Analytics" ? '#1e3a8a' : 'white', color: activeView === "Data Analytics" ? 'white' : '#475569', boxShadow: activeView === "Data Analytics" ? '0 4px 8px rgba(30, 58, 138, 0.25)' : '0 1px 3px rgba(0,0,0,0.05)' }}>DATA ANALYTICS</button>
          </div>

          <div style={{ marginTop: 'auto', background: 'white', borderRadius: '16px', border: '1px solid #cbd5e1', overflow: 'hidden', pointerEvents: 'auto' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#1e3a8a', color: 'white', padding: '14px 15px' }}>
              <button style={{ background: 'none', border: 'none', color: 'white', cursor: 'pointer', fontSize: '18px', fontWeight: 'bold', padding: '0 5px', pointerEvents: 'auto' }} onClick={() => { setCalCurrentDate(new Date(calCurrentDate.getFullYear(), calCurrentDate.getMonth() - 1)); setShowMonthPicker(false); }}>&#9664;</button>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <div onClick={() => setShowMonthPicker(!showMonthPicker)} style={{ fontWeight: 'bold', fontSize: '14px', cursor: 'pointer', letterSpacing: '0.5px', pointerEvents: 'auto' }}>{calCurrentDate.toLocaleString('default', { month: 'long' }).toUpperCase()} {calCurrentDate.getFullYear()}</div>
              </div>
              <button style={{ background: 'none', border: 'none', color: 'white', cursor: 'pointer', fontSize: '18px', fontWeight: 'bold', padding: '0 5px', pointerEvents: 'auto' }} onClick={() => { setCalCurrentDate(new Date(calCurrentDate.getFullYear(), calCurrentDate.getMonth() + 1)); setShowMonthPicker(false); }}>&#9654;</button>
            </div>
            
            <div style={{ padding: '15px', minHeight: '320px', display: 'flex', flexDirection: 'column', background: '#eef2f6', pointerEvents: 'auto' }}>
              {showMonthPicker ? (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  <input type="number" className="input-ui" style={{ textAlign: 'center', fontWeight: 'bold' }} value={calCurrentDate.getFullYear()} onChange={(e) => setCalCurrentDate(new Date(parseInt(e.target.value), calCurrentDate.getMonth()))} />
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '5px' }}>
                    {Array.from({length: 12}).map((_, i) => {
                      const d = new Date(2000, i); const isSelected = i === calCurrentDate.getMonth();
                      return ( <button key={i} onClick={() => { setCalCurrentDate(new Date(calCurrentDate.getFullYear(), i)); setShowMonthPicker(false); }} style={{ padding: '8px 4px', fontSize: '11px', fontWeight: 'bold', borderRadius: '6px', border: 'none', cursor: 'pointer', background: isSelected ? '#f97316' : 'white', color: isSelected ? 'white' : '#475569', fontFamily: 'Poppins', boxShadow: '0 1px 2px rgba(0,0,0,0.05)', pointerEvents: 'auto' }}>{d.toLocaleString('default', { month: 'short' })}</button> );
                    })}
                  </div>
                </div>
              ) : (
                <div style={{ height: 'auto', marginBottom: '15px' }}>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '4px', textAlign: 'center', fontSize: '11px', fontWeight: '800', marginBottom: '8px' }}>
                    {['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'].map((d, i) => <div key={d} style={{ color: i === 0 ? '#f97316' : '#64748b' }}>{d}</div>)}
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gridTemplateRows: 'repeat(6, 32px)', gap: '4px', textAlign: 'center', fontSize: '12px' }}>
                    {renderCalendarDays()}
                  </div>
                </div>
              )}

              <button style={{ width: '100%', marginTop: 'auto', padding: '10px', fontSize: '12px', fontWeight: '800', background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px', cursor: 'pointer', color: '#f97316', transition: 'all 0.2s', fontFamily: 'Poppins', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', pointerEvents: 'auto' }} 
                      onMouseOver={e => { e.currentTarget.style.background = '#f97316'; e.currentTarget.style.color = 'white'; e.currentTarget.style.borderColor = '#ea580c'; }} 
                      onMouseOut={e => { e.currentTarget.style.background = 'white'; e.currentTarget.style.color = '#f97316'; e.currentTarget.style.borderColor = '#cbd5e1'; }}
                      onClick={() => { setDateRange({start: null, end: null}); setCurrentPage(1); }}>
                Reset Filter Tanggal
              </button>
            </div>
          </div>
        </aside>

        {/* ================= 3. MAIN AREA ================= */}
        {activeView === "Data Input" && (
          <>
            {/* FORCE Z-INDEX & POINTER EVENTS */}
            <section className="input-section custom-scroll page-transition" style={{ width: '420px', background: '#eef2f6', padding: '25px', borderRadius: '16px', border: '1px solid #cbd5e1', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '18px', flexShrink: 0, position: 'relative', zIndex: 9999, pointerEvents: 'auto' }}> 
              
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Alert Type</label>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button style={{ flex: 1, padding: '10px', borderRadius: '8px', border: '1px solid #cbd5e1', fontWeight: '700', fontSize: '13px', transition: 'all 0.2s', background: alertTypeBtn === "Product" ? '#1e3a8a' : 'white', color: alertTypeBtn === "Product" ? 'white' : '#64748b', cursor: 'pointer', boxShadow: alertTypeBtn === "Product" ? '0 2px 4px rgba(0,0,0,0.1)' : '0 1px 2px rgba(0,0,0,0.05)', fontFamily: 'Poppins', pointerEvents: 'auto' }} onClick={() => setAlertTypeBtn("Product")}>Product</button>
                  <button style={{ flex: 1, padding: '10px', borderRadius: '8px', border: '1px solid #cbd5e1', fontWeight: '700', fontSize: '13px', transition: 'all 0.2s', background: alertTypeBtn === "Process" ? '#1e3a8a' : 'white', color: alertTypeBtn === "Process" ? 'white' : '#64748b', cursor: 'pointer', boxShadow: alertTypeBtn === "Process" ? '0 2px 4px rgba(0,0,0,0.1)' : '0 1px 2px rgba(0,0,0,0.05)', fontFamily: 'Poppins', pointerEvents: 'auto' }} onClick={() => setAlertTypeBtn("Process")}>Process</button>
                </div>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Type of Alert</label>
                <select className="input-ui" value={typeAlertSelect} onChange={handleTypeAlertChange}>
                  <option value="">-- Pilih Jenis Alert --</option>
                  {Object.keys(alertCategoryMapping).map((opt, idx) => (
                    <option key={idx} value={opt}>{opt}</option>
                  ))}
                </select>
              </div>

              {/* FITUR 1: DROPDOWN OTHER SEARCHABLE */}
              {typeAlertSelect === "Other" && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  <label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Kategori Other</label>
                  <div style={{ position: 'relative' }}>
                    <input 
                      className="input-ui" 
                      value={otherSelect} 
                      onChange={(e) => { setOtherSelect(e.target.value); setShowOtherDropdown(true); }} 
                      onFocus={() => setShowOtherDropdown(true)}
                      onBlur={() => setTimeout(() => setShowOtherDropdown(false), 200)}
                      placeholder="-- Ketik atau pilih kategori --" 
                    />
                    {showOtherDropdown && filteredOther.length > 0 && (
                      <ul className="custom-scroll" style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px', maxHeight: '150px', overflowY: 'auto', zIndex: 100, padding: 0, margin: '4px 0 0 0', listStyle: 'none', boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
                        {filteredOther.map((opt, idx) => (
                          <li key={idx} style={{ padding: '8px 12px', cursor: 'pointer', fontSize: '12px', borderBottom: '1px solid #f1f5f9' }} 
                              onMouseDown={() => { setOtherSelect(opt); setShowOtherDropdown(false); }}>
                            {opt}
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              )}
              
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>
                   Initiator yang melapor
                </label>
                <div style={{ position: 'relative' }}>
                  <input 
                    className="input-ui" 
                    value={initiator} 
                    onChange={(e) => { setInitiator(e.target.value); setShowInitiatorDropdown(true); }} 
                    onFocus={() => setShowInitiatorDropdown(true)}
                    onBlur={() => setTimeout(() => setShowInitiatorDropdown(false), 200)}
                    placeholder="-- Ketik nama atau klik panah --" 
                  />
                  {showInitiatorDropdown && filteredInitiators.length > 0 && (
                    <ul className="custom-scroll" style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px', maxHeight: '150px', overflowY: 'auto', zIndex: 100, padding: 0, margin: '4px 0 0 0', listStyle: 'none', boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
                      {filteredInitiators.map((opt, idx) => (
                        <li key={idx} style={{ padding: '8px 12px', cursor: 'pointer', fontSize: '12px', borderBottom: '1px solid #f1f5f9' }} 
                            onMouseDown={() => { setInitiator(opt); setShowInitiatorDropdown(false); }}>
                          {opt}
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Shift</label>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px' }}>
                  {[1, 2, 3].map(s => (
                    <button key={s} style={{ padding: '8px', borderRadius: '8px', border: '1px solid #cbd5e1', fontWeight: 'bold', background: shiftBtn === s.toString() ? '#1e3a8a' : 'white', color: shiftBtn === s.toString() ? 'white' : '#64748b', cursor: 'pointer', transition: 'all 0.2s', boxShadow: shiftBtn === s.toString() ? '0 2px 4px rgba(59, 130, 246, 0.3)' : '0 1px 2px rgba(0,0,0,0.05)', fontFamily: 'Poppins', pointerEvents: 'auto' }} onClick={() => setShiftBtn(s.toString())}>{s}</button>
                  ))}
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '15px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}><label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Tanggal Temuan</label><input className="input-ui" type="date" value={tglTemuan} onChange={e => setTglTemuan(e.target.value)} /></div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}><label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Jam</label><input className="input-ui" type="time" value={wktTemuan} onChange={e => setWktTemuan(e.target.value)} /></div>
              </div>
              
              <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '15px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}><label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Tanggal Correction</label><input className="input-ui" type="date" value={tglCorr} onChange={e => setTglCorr(e.target.value)} /></div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}><label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Jam</label><input className="input-ui" type="time" value={wktCorr} onChange={e => setWktCorr(e.target.value)} /></div>
              </div>
              
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '15px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}><label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>MO Number</label><input className="input-ui" type="text" placeholder="MO-XXX" value={moNum} onChange={e => setMoNum(e.target.value)} /></div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}><label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Item Code</label><input className="input-ui" type="text" placeholder="IT-XXX" value={itemCode} onChange={e => setItemCode(e.target.value)} /></div>
              </div>
              
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Kode Mesin & Team</label>
                <div style={{ display: 'flex', gap: '10px' }}>
                  <input className="input-ui" style={{flex: 1}} type="text" placeholder="Cth: IM01" value={inputMesin} onChange={handleMesinChange} />
                  <div style={{ flex: 1, padding: '10px', borderRadius: '8px', backgroundColor: teamInfo.bg, color: teamInfo.color, fontWeight: '700', fontSize: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', border: `1px solid ${teamInfo.border}` }}>{teamInfo.name}</div>
                </div>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}><label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Event Description</label><textarea className="input-ui custom-scroll" style={{ minHeight: '70px', resize: 'vertical' }} placeholder="Ceritakan detail masalah..." value={eventDesc} onChange={e => setEventDesc(e.target.value)} /></div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}><label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Correction Action</label><textarea className="input-ui custom-scroll" style={{ minHeight: '70px', resize: 'vertical' }} placeholder="Tindakan yang dilakukan..." value={corrAction} onChange={e => setCorrAction(e.target.value)} /></div>
              
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '13px', fontWeight: '800', color: '#1e293b' }}>Action by Operator</label>
                <div style={{ position: 'relative' }}>
                  <input 
                    className="input-ui" 
                    value={actionOp} 
                    onChange={(e) => { setActionOp(e.target.value); setShowActionOpDropdown(true); }} 
                    onFocus={() => setShowActionOpDropdown(true)}
                    onBlur={() => setTimeout(() => setShowActionOpDropdown(false), 200)}
                    placeholder="-- Ketik nama atau klik panah --" 
                  />
                  {showActionOpDropdown && filteredActionOps.length > 0 && (
                    <ul className="custom-scroll" style={{ position: 'absolute', bottom: '100%', left: 0, right: 0, background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px', maxHeight: '150px', overflowY: 'auto', zIndex: 100, padding: 0, margin: '0 0 4px 0', listStyle: 'none', boxShadow: '0 -4px 6px rgba(0,0,0,0.1)' }}>
                      {filteredActionOps.map((opt, idx) => (
                        <li key={idx} style={{ padding: '8px 12px', cursor: 'pointer', fontSize: '12px', borderBottom: '1px solid #f1f5f9' }} 
                            onMouseDown={() => { setActionOp(opt); setShowActionOpDropdown(false); }}>
                          {opt}
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>

              <button className="btn-simpan" style={{ padding: '14px', borderRadius: '10px', border: 'none', fontWeight: 'bold', fontSize: '14px', cursor: 'pointer', marginTop: '10px', flexShrink: 0 }} onClick={submitData}>+ Simpan Data Alert</button>
            </section>

            {/* AREA LOG KANAN */}
            <section className="print-area-container page-transition" style={{ flex: 1, background: 'transparent', display: 'flex', flexDirection: 'column', overflow: 'hidden', minHeight: 0, minWidth: '700px', position: 'relative', zIndex: 9999, pointerEvents: 'auto' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', padding: '0 5px', flexShrink: 0, background: 'white', padding: '15px 20px', borderRadius: '16px', boxShadow: '0 4px 6px rgba(0,0,0,0.04)' }}>
                <div style={{ fontSize: '20px', fontWeight: '900', color: '#0f172a', letterSpacing: '-0.5px', display: 'flex', alignItems: 'center', fontFamily: 'system-ui, -apple-system, sans-serif' }}>
                  Database Logs 
                  {dateRange.start && ( <span style={{ fontSize: "11px", color: "#c2410c", background: "#ffedd5", padding: "4px 12px", borderRadius: "20px", marginLeft: "15px", verticalAlign: "middle", border: '1px solid #fed7aa', fontFamily: "'Poppins', sans-serif" }}> 🗓 Rentang: {dateRange.start} {dateRange.end ? ` - ${dateRange.end}` : ''} </span> )}
                </div>
                
                <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                   <button onClick={handleExportExcel} className="btn-export">
                     <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                     Unduh Excel
                   </button>
                </div>
              </div>
              
              <div className="custom-scroll print-scroll-reset" style={{ flex: 1, overflowY: 'auto', paddingRight: '10px', minHeight: 0 }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}> 
                      <div style={{ display: 'flex', padding: '0 20px', color: '#64748b', fontWeight: '800', fontSize: '11px', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: '4px', flexShrink: 0 }}>
                        <div style={{ width: '16%' }}>Waktu Temuan</div><div style={{ width: '22%' }}>Target</div><div style={{ width: '16%' }}>Personel</div><div style={{ width: '26%' }}>Informasi Kasus</div><div style={{ width: '20%' }}>Tindakan & Waktu</div>
                      </div>

                      {pageData.length === 0 ? (
                        <div style={{textAlign:"center", padding:"50px", background: 'white', borderRadius: '16px', color: '#94a3b8', border: '1px dashed #cbd5e1', fontWeight: '600'}}>Belum ada log data pada rentang waktu ini.</div>
                      ) : (
                        pageData.map((row) => (
                          <div key={row.id} className="card-row" style={{ display: 'flex', padding: '16px 20px', paddingRight: '70px', alignItems: 'flex-start', position: 'relative' }}>
                            <button onClick={() => handleDelete(row.id)} style={{ position: 'absolute', top: '16px', right: '16px', background: '#fee2e2', color: '#ef4444', border: '1px solid #fecaca', borderRadius: '6px', padding: '5px 10px', fontSize: '10px', fontWeight: 'bold', cursor: 'pointer', transition: 'all 0.2s', fontFamily: 'Poppins', pointerEvents: 'auto' }} onMouseOver={(e) => { e.currentTarget.style.background = '#ef4444'; e.currentTarget.style.color = 'white'; }} onMouseOut={(e) => { e.currentTarget.style.background = '#fee2e2'; e.currentTarget.style.color = '#ef4444'; }}>Hapus</button>

                            <div style={{ width: '16%', display: 'flex', flexDirection: 'column', gap: '6px', minWidth: 0, paddingRight: '10px' }}>
                              <div style={{fontWeight: "800", color: "#0f172a", fontSize: '13px', wordWrap: 'break-word'}}>{row.tglT}</div>
                              <div style={{color: "#f97316", fontWeight: "600", fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px'}}>
                                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{flexShrink: 0}}><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg><span style={{wordWrap: 'break-word'}}>{row.wktT}</span>
                              </div>
                              <div style={{ marginTop: '4px' }}><span style={{ fontSize: '9px', fontWeight: '800', background: row.alertType === 'Product' ? '#dbeafe' : '#ffedd5', color: row.alertType === 'Product' ? '#1e40af' : '#c2410c', padding: '4px 10px', borderRadius: '12px', letterSpacing: '0.5px', wordWrap: 'break-word', display: 'inline-block' }}>{String(row.alertType || '').toUpperCase()}</span></div>
                            </div>
                            
                            <div style={{ width: '22%', display: 'flex', flexDirection: 'column', gap: '8px', minWidth: 0, paddingRight: '10px' }}>
                              <div style={{display: 'flex', gap: '6px', flexWrap: 'wrap'}}>
                                <span style={{color: "#475569", fontWeight: "600", fontSize: '11px', background: '#f1f5f9', padding: '3px 8px', borderRadius: '6px', border: '1px solid #e2e8f0', wordWrap: 'break-word'}}>MO: {row.mo}</span><span style={{color: "#475569", fontWeight: "600", fontSize: '11px', background: '#f1f5f9', padding: '3px 8px', borderRadius: '6px', border: '1px solid #e2e8f0', wordWrap: 'break-word'}}>IT: {row.item}</span>
                              </div>
                              <div style={{display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap'}}>
                                <span style={{fontWeight: '900', color: '#1e3a8a', fontSize: '15px', wordWrap: 'break-word'}}>{row.msn}</span><span style={{fontSize: '9px', fontWeight: '800', background: '#e2e8f0', color: '#475569', padding: '3px 8px', borderRadius: '6px', letterSpacing: '0.5px', wordWrap: 'break-word'}}>{row.tm}</span>
                              </div>
                            </div>
                            
                            <div style={{ width: '16%', display: 'flex', flexDirection: 'column', gap: '6px', fontSize: '11px', color: '#334155', minWidth: 0, paddingRight: '10px' }}>
                              <div style={{display: 'flex', justifyContent: 'space-between'}}><span style={{color:'#94a3b8', flexShrink: 0}}>Shift</span><span style={{fontWeight:'800', color:'#0f172a', wordWrap: 'break-word', textAlign: 'right'}}>{row.shift}</span></div>
                              <div style={{display: 'flex', justifyContent: 'space-between'}}><span style={{color:'#94a3b8', flexShrink: 0}}>Init</span><span style={{fontWeight:'600', wordWrap: 'break-word', textAlign: 'right'}}>{row.opr}</span></div>
                              <div style={{display: 'flex', justifyContent: 'space-between'}}><span style={{color:'#94a3b8', flexShrink: 0}}>Oper</span><span style={{fontWeight:'600', wordWrap: 'break-word', textAlign: 'right'}}>{row.actOp}</span></div>
                            </div>
                            
                            <div style={{ width: '26%', display: 'flex', flexDirection: 'column', gap: '4px', paddingRight: '15px', minWidth: 0 }}>
                              <div style={{fontWeight:"800", color:"#dc2626", fontSize: '14px', wordWrap: 'break-word'}}>{row.alert}</div><div style={{color:"#475569", fontSize:"12px", lineHeight: '1.5', wordWrap: 'break-word'}}>{row.desc}</div>
                            </div>

                            <div style={{ width: '20%', display: 'flex', flexDirection: 'column', gap: '6px', minWidth: 0 }}>
                              <div style={{color:"#059669", fontWeight: "700", fontSize:"13px", lineHeight: '1.4', display: 'flex', alignItems: 'flex-start', gap: '4px', wordWrap: 'break-word'}}>
                                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" style={{marginTop:'3px', flexShrink: 0}}><polyline points="20 6 9 17 4 12"></polyline></svg><span style={{wordWrap: 'break-word'}}>{row.act}</span>
                              </div>
                              {row.tglC && row.tglC !== '-' && ( <div style={{fontSize: '11px', color: '#64748b', display: 'flex', alignItems: 'flex-start', gap: '4px', background: '#f1f5f9', padding: '4px 8px', borderRadius: '6px', width: 'fit-content'}}> <span style={{flexShrink: 0}}>Done:</span><span style={{fontWeight: '700', wordWrap: 'break-word'}}>{row.tglC}</span><span style={{color: '#94a3b8', wordWrap: 'break-word'}}>({row.wktC})</span></div> )}
                            </div>
                          </div>
                        ))
                      )}
                </div>
              </div>

              {filteredData.length > 0 && (
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', marginTop: '15px', gap: '15px', padding: '12px 0', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', flexShrink: 0, boxShadow: '0 4px 6px rgba(0,0,0,0.04)' }}>
                  <button disabled={currentPage === 1} onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))} style={{ padding: '8px 16px', borderRadius: '6px', border: '1px solid #cbd5e1', background: currentPage === 1 ? '#f1f5f9' : 'white', color: currentPage === 1 ? '#94a3b8' : '#334155', cursor: currentPage === 1 ? 'not-allowed' : 'pointer', fontWeight: 'bold', fontSize: '13px', transition: 'all 0.2s', fontFamily: 'Poppins', pointerEvents: 'auto' }}>&#9664; Prev</button>
                  <span style={{ fontSize: '14px', fontWeight: '800', color: '#1e3a8a' }}>Halaman {currentPage} dari {totalPages}</span>
                  <button disabled={currentPage === totalPages} onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))} style={{ padding: '8px 16px', borderRadius: '6px', border: '1px solid #cbd5e1', background: currentPage === totalPages ? '#f1f5f9' : 'white', color: currentPage === totalPages ? '#94a3b8' : '#334155', cursor: currentPage === totalPages ? 'not-allowed' : 'pointer', fontWeight: 'bold', fontSize: '13px', transition: 'all 0.2s', fontFamily: 'Poppins', pointerEvents: 'auto' }}>Next &#9654;</button>
                </div>
              )}

            </section>
          </>
        )}

        {/* ================= 4. DATA MANAGEMENT (PIVOT) ================= */}
        {activeView === "Data Management" && (
          <section className="print-area-container page-transition custom-scroll" style={{ flex: 1, minWidth: '700px', background: 'rgba(255,255,255,0.95)', backdropFilter: 'blur(10px)', borderRadius: '16px', border: '1px solid #e2e8f0', padding: '25px', display: 'flex', flexDirection: 'column', boxShadow: '0 4px 12px rgba(0,0,0,0.05)', overflowY: 'auto', position: 'relative', zIndex: 9999, pointerEvents: 'auto' }}>
            <h2 style={{ color: '#0f172a', margin: '0 0 20px 0', fontSize: '24px', fontWeight: '900', flexShrink: 0, fontFamily: 'system-ui, -apple-system, sans-serif' }}>Data Management Pivot</h2>
            
            <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', height: 'max-content' }}>
              {[ { title: 'Capsule Team', data: capsuleSummary }, { title: 'Mono Team', data: monoSummary }, { title: 'ND3H / D3H', data: nd3hSummary } ].map((table, i) => (
                <div key={i} style={{ flex: 1, minWidth: '280px', background: 'white', borderRadius: '16px', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', overflow: 'hidden', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                  <div style={{ background: '#1e3a8a', padding: '18px', fontWeight: '800', color: 'white', textAlign: 'center', letterSpacing: '0.5px', flexShrink: 0, fontFamily: 'system-ui, -apple-system, sans-serif', fontSize: '16px' }}>{table.title}</div>
                  
                  <div style={{ padding: '15px 20px 25px 20px', display: 'flex', flexDirection: 'column' }}>
                    {table.data.length === 0 ? <div style={{textAlign:'center', color:'#94a3b8', padding:'30px', fontWeight: '600'}}>No data</div> : 
                      table.data.map((r, x) => (
                        <div key={x} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '14px 0', borderBottom: '1px solid #e2e8f0', fontSize: '14px' }}>
                          <span style={{color: '#334155', fontWeight: '600'}}>{r.label}</span>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexShrink: 0 }}>
                             <span style={{fontWeight: '900', color: '#f97316', fontSize: '16px'}}>{r.count}</span>
                             <span style={{fontSize:'11px', fontWeight: '800', background: '#ffedd5', color: '#ea580c', padding: '3px 8px', borderRadius: '6px'}}>{r.percent}</span>
                          </div>
                        </div>
                      ))
                    }
                  </div>
                </div>
              ))}
            </div>
          </section>
        )}

        {/* ================= 5. DATA ANALYTICS ================= */}
        {activeView === "Data Analytics" && (
          <section className="print-area-container page-transition" style={{ flex: 1, minWidth: '700px', background: 'rgba(255,255,255,0.95)', backdropFilter: 'blur(10px)', borderRadius: '16px', border: '1px solid #e2e8f0', padding: '25px', display: 'flex', flexDirection: 'column', boxShadow: '0 4px 12px rgba(0,0,0,0.05)', overflow: 'hidden', position: 'relative', zIndex: 9999, pointerEvents: 'auto' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexShrink: 0 }}>
               <h2 style={{ color: '#0f172a', margin: '0', fontSize: '24px', fontWeight: '900', fontFamily: 'system-ui, -apple-system, sans-serif' }}>Analytics Dashboard</h2>
               {dateRange.start && ( <span style={{ fontSize: "13px", color: "#c2410c", background: "#ffedd5", padding: "6px 16px", borderRadius: "20px", fontWeight: "800" }}> 🗓 Rentang Filter: {dateRange.start} {dateRange.end ? ` - ${dateRange.end}` : ''} </span> )}
            </div>
            <div className="custom-scroll print-scroll-reset" style={{ display: 'flex', flexWrap: 'wrap', gap: '20px', flex: 1, overflowY: 'auto', paddingBottom: '10px', minHeight: 0 }}>
              <ParetoChart title="Capsule" data={capsuleSummary} />
              <ParetoChart title="Mono" data={monoSummary} />
              <ParetoChart title="ND3H / D3H" data={nd3hSummary} />
            </div>
          </section>
        )}

      </div>
    </div>
  );
}
#!/usr/bin/python3
# -*- coding: utf-8 -*-

import sys
import codecs
import csv
import re
import datetime
import copy

#=============================================================================
# NY補習校W校 図書システム (School Pro) 利用者情報 年度末更新プログラム
# 
# 年度の始めに学校から提供される生徒名簿を元に、
# School Pro の利用者情報を更新する
# 幼児部から高等部まですべての利用者を一括で更新する
#
# ・既にSchool Pro に登録されている「生徒」「生徒（中高）」「保護者」
# 「保護者（中高）」の利用者は新年度も在籍している場合は学年、クラスを更新する
# 　「学部」も更新(例:幼年長から初等1年に進学した場合は学部を「初等部」に変更
#   必要なら利用者区分も変更する(初等6年から中等1年に進学した場合は生徒カード
# 　の利用者区分を「生徒」から「生徒（中高）」に変更する。また、保護者カードは
# 　「保護者」から「保護者（中高）」に変更する)
#
# ・既にSchool Pro に「生徒」「生徒（中高）」「保護者」「保護者（中高）」
# 　として登録されていたが、新年度は在籍していない利用者は利用者区分を
# 「卒業生・退学者」に変更する。生徒と保護者は区別しない
# 
# ・今までSchool Pro に登録されていなかったが新年度名簿に入っている利用者は
#   新しい利用者IDを割り当て新規登録する。一人の生徒につき生徒カードと保護者
# 　カードの2人分の利用者情報を登録する。生徒カードの利用者区分は「生徒」又は
#  「生徒（中高）」となる。保護者カードは「保護者」もしくは「保護者（中高）」
# 　となる。新規利用者はカード印刷が必要。
#   
#=============================================================================
# 使い方: 
# update_user.py  <OUT School Pro 利用者情報CSV> <OUT 確認用デバッグ出力CSV> 
#           <IN 現在のSchool Pro利用者情報CSV> <IN 学校提供 新年度生徒名簿CSV>
#
#  <OUT School Pro 利用者情報CSV> :
#　　 School Pro 利用者情報CSV出力、UTF-8 (BOM無しなので注意)。
#     ＊＊ nkfでSJISに戻してから＊＊ School Proにインポートする。
#     各行のフォーマットは後述の "IN 現在のSchool Pro利用者情報CSV" と同様
#     だが、列 8 (I) 利用者アルファベット名のみは注意が必要。
#     School Pro からエクスポートされる時点では姓と名の間のコンマが空白に置換
#  　 されているが、インポートする際には*全角の*コンマ「，」を姓と名の間に
#     入れておく。この全角コンマはインポートするとSchool Pro により半角コンマ
#     に変更される。
#    
#  <OUT 確認用デバッグ出力CSV> :
#     処理内容確認のためのデバッグ出力。UTF-8 (BOM無し!)
#     このファイルは年度更新処理には不要だが、一応中身を確認した方が良い。
#     既に登録されている利用者情報に何らかの変更が加わった場合は「変更」と
#     ラベルされた行を出力。旧データと更新された部分を2行使って示す。
#     新規登録された利用者情報は「新規」とラベルされた行を出力する。1行に
#     その利用者の情報を出力する。
#
#  <IN 現在のSchool Pro利用者情報CSV> :
#　　 School Pro からエクスポートした現在の利用者情報CSV。UTF-8 (BOM無し!)
#     School Pro から出力された時点ではCSVは SJISフォーマットなので、必ず
#     nkf 等を使い UTF-8に変更すること。BOMを付けると codesc がバグるので
#     入れないこと。nkf の場合は -w80 とすれば BOM無しのUTF-8出力となる。
#     ＊＊＊ 1行目は無視されるので注意 ＊＊＊
#     各行、以下のフォーマットで図書利用者一人分の利用者情報を示す。
#     一般に、生徒一人につき生徒と保護者の二人分(2行分)のデータが存在
#       列 0  (A) : 利用者番号 例 20201999
#       列 1  (B) : 利用者区分 例 '生徒'
#       列 2  (C) : 学科 例 '初等部'
#       列 3  (D) : 学年 例 1
#       列 4  (E) : クラス 例 4
#       列 5  (F) : 出席番号(便宜的につけたもの、無くてもよい) 例 12
#       列 6  (G) : 性別 例 '女'
#       列 7  (H) : 利用者氏名 例 '北条 政子'
#       列 8  (I) : フリガナ(実際にはアルファベット表記) 例 'HOJO MASAKO'
#                   School Pro上では姓と名の間にコンマがあるが、CSV exportの
#                   際に削除され空白に変更される。
#       列 9  (J) : 郵便番号 : 特に使用しない。空白にしておく
#       列 10 (K) : 保護者氏名 例 '北条 時政'
#       列 11 (L) : 保護者住所(実際は E-mail) 例 'hojo_tokimasa@gmail.com'
#       列 12 (M) : 入学年　例: 2020
#       列 13 (N) : 任意集計項目: メモなど任意の文字列、登録年月など
#       列 14 (O) : 転退 : 使用しない、空白で良い
#       列 15 (P) : 転退日 : 使用しない、空白で良い
#       列 16 (Q) : 留 : 使用しない、空白で良い
#
#  <IN 学校提供 新年度生徒名簿CSV> :
#     学校から提供された今年度の利用者情報CSV
#     ＊＊＊ 1行目は無視されるので注意 ＊＊＊
#     各行、以下のフォーマットで新年度在籍生徒一人分の利用者情報を示す。
#       列 0  (A) : 学籍番号 例 'T201999'
#       列 1  (B) : アルファベット生徒名 例: 'MARCH, JOSEPHINE'
#       列 2  (C) : 性別英語 例 'F', 'M'
#       列 3  (D) : 日本語生徒名 例 '北条　政子'
#       列 4  (E) : アルファベット保護者名 : 使用しないので空白でも良い
#       列 5  (F) : 日本語保護者名 例 '北条　時政'
#       列 6  (G) : 課程区分(学籍番号の先頭アルファベット) 'A' か 'T'
#       列 7  (H) : 学年 例 1
#       列 8  (I) : クラス 例 4
#       列 9  (J) : 保護者電話番号 : 使用しないので空白でも良い
#       列 10 (K) : 保護者 E-mail 例 'hojo_tokimasa@gmail.com'
#
#-----------------------------------------------------------------------------
# メモ: 入出力ファイルの文字コード変換。
# WSL のSJISロケールは一部の文字コードを正しく処理できないバグがあるうえ、
# SJIS そのものがASCII文字とのコード重複など問題が多い。そのため プログラムで
# 処理するときは全て UTF-8で行う。School Pro は SJIS しかサポートしていない
# ので、以下の通りコード変換を行う。
#
# 入力CSV を UTF-8に変換 
#       (BOMを付けないため -w8 ではなく -w80 を指定するのに注意)
#
#     nkf -c  -w80 users2019.csv
#
# 出力CSV を SJISに戻す方法
#
#     nkf -s -c utf8out.txt
#
#-----------------------------------------------------------------------------
# 最終的に出力する School Pro 利用者CSVの行フォーマット:
#
#   0                             5                            
# [ id, cat, sect, grade, class, cno, sx, name, ename, postal,
#   10                              15
#  gname, gaddr, year, arbit, grad, grad_date, sty ]

#*****************************************************************************
# グローバル変数
#*****************************************************************************
g_file = '' # 現在読んでいるファイル名
g_line = 0  # 現在の行

g_debug = 0 # デバッグフラグ

#-----------------------------------------------------------------------------
# 辞書: 利用者情報を格納
# キー : School Pro 利用者ID
# 値 : ToshoUser のオブジェクト、利用者情報を保持

current_users = {}   # 既に School Pro に登録されていた利用者情報
updated_users = {}   # 新年度向けに更新された利用者情報

#-----------------------------------------------------------------------------
# 辞書: 各クラスに属する利用者IDのリストを保持
# Key = "P-1", "P-2", "P-3", "0-1", "0-2", "0-3", ...., 
#       "9-1", "9-2", "9-3", "A-1", "A-2", "A-3", "B-1", "B-2", "B-3"
# elem = List of tuple(Student UID, Guardian UID)
# e.g. 
# cllist["P-1"] = [ (student0_uid, parent0_uid), 
#                   (student1_uid, parent1_uid), ..., ]

cllist = {}

#*****************************************************************************
# サブルーチン
#*****************************************************************************

#=============================================================================
# 標準エラー出力にメッセージを出す

def print_stderr(*x):
	print ("ファイル", g_file, "行", g_line, ":", *x, file=sys.stderr)

def print_stderr_exit(*x):
	print ("ファイル", g_file, "行", g_line, ":", *x, file=sys.stderr)
	sys.exit(1)


#=============================================================================
# 日付文字列

def date_str():
	
	# 今日の日付
	dt = datetime.date.today()
	cy = dt.year
	cm = dt.month
	cd = dt.day

	return  "%04u%02u%02u" % (cy, cm, cd) 

#=============================================================================
# CSVの差分出力。x と y が異なれば x, 同じなら空文字列を返す
# ただし差分ありの場合にx が空文字列だと差分がない場合と見分けが付かないので
# 明示的に'(空白)’と表示

def take_diff(x,y):
	if (x == y):    # 差分無し
		return ''
	return '（空白）' if x == '' else x # 差分あり xが空白なら（空白）と表示

#=============================================================================
# 文字列の両側の空白を削除する

def rmsp(s):
	
	while True:
		# 左側空白を削除
		result = re.subn(r'$[\s　]+', '', s)
		
		if result[1]:
			s = result[0]
			continue
		
		# 右側
		result = re.subn(r'[\s　]+$', '', s)
		if result[1]:
			s = result[0]
			continue
		
		break
	
	return s

#=============================================================================
# 利用者名の半角文字を全角に変更

kana1  = r'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｧｨｩｪｫｯｬｭｮｰ()'
wkana1 = r'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホ' \
	     r'マミムメモヤユヨラリルレロワヲンァィゥェォッャュョー（）'

kana2  = r'ｶﾞｷﾞｸﾞｹﾞｺﾞｻﾞｼﾞｽﾞｾﾞｿﾞﾀﾞﾁﾞﾂﾞﾃﾞﾄﾞﾊﾞﾋﾞﾌﾞﾍﾞﾎﾞﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟｳﾞ'
wkana2 = r'ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポヴ'

def conv_1bto2b(s):
	# conv 1B alphabet and kana into 2B
	o = ''
	
	table = str.maketrans(kana1, wkana1)
	s2 = s.translate(table)

	idx = 0	
	l = len(s2)

	while idx < l-1:
		c2 = s[idx:idx+2]
		x = kana2.find(c2)

		if x >= 0:
			o = o + wkana2[x >> 1]
			idx += 2
			continue

		o = o + s2[idx]
		idx += 1

	if idx < l:
		o = o + s2[idx]

	o = re.sub(r'([' + wkana1 + wkana2 + r'])-', r'\1ー', o)

	if g_debug and s != o:
		print ("conv_1bto2b |%s|->|%s|" % (s, o))

	return o

#=============================================================================
# 名前の空白などの書式を標準形に統一する

def normalize_name(name):

	o = name
	name = rmsp(name) # 両端の余分な空白を除去
	name = conv_1bto2b(name)  # 半角カナを全角に変換

	name = re.sub(r'[\s　]+', '　', name) # 空白は全角に。複数空白は１つに。

	# 保護者の"（保）" の前にある空白のみは半角にする

	result = re.subn(r'[\s　]*（保）', '', name)
	if result[1]: # '（保）'とマッチした場合
		name = result[0]  # （保）の前の部分を取り出し
		name = rmsp(name) # 両側の空白を除き
		name = name + " （保）" # 改めて半角空白と（保）を追加

	if g_debug and o != name:
		print ("normalize_name |%s|->|%s|" % (o, name))
	return name

#=============================================================================
# 利用者英名中の全角文字を半角に統一する

walpha = 'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ' \
		 'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
aalpha = "ABCDEFGHIJKLMNOPRQSTUVWXYZabcdefghijklmnoprqstuvwxyz"

def alpha_ascii(s):

	o = ''
	idx = 0
	l = len(s)

	while idx < l:

		c = s[idx]
		x = walpha.find(c) # 全角文字を探す
			
		if (x >= 0): # 全角文字が見つかった場合
			o = o + aalpha[x] # 対応する半角文字に変換
			idx += 1
			continue

		# No match, just add the character
		o = o + s[idx]
		idx += 1

	return o

#=============================================================================
# 英名の姓と名の間のコンマを復活させる (School Pro がCSV出力時に除去)
# School Pro はインポートの際に半角コンマは正しく受け付けない。
# 全角のコンマは受け付けられ半角のコンマに変換されるので全角にする。

def recover_comma(s):
	s = re.sub(r'\s(\S+\s*)$', r'，\1', s)	# 最後の空白を全角コンマに置き換え
	return s

#=============================================================================
# 英名を正規化

def adjust_ename(s):

	orig = s
	s = rmsp(s) # 空白除去
	s = alpha_ascii(s) # アルファベットを半角に統一
	s = re.sub(r'\'', r'`', s)	# quote を back quote に変換
	s = re.sub(r'\,', r'，', s)	# コンマが既に付いていれば全角に変換
	s = re.sub(r'[\s　]*，[\s　]*', r'，', s) # コンマの前後の空白を除去	
	s = re.sub(r'[\s　]+', ' ', s)	# 複数の空白は1個に減らす
	s = re.sub(r'`[\s　]+', '`', s)	# back quote の後のスペースを削除
	s = s.upper() # 全て大文字に統一

	if g_debug and orig != s:
		print ("adjust_ename |%s|->|%s|" % (orig, s))

	return s


#=============================================================================
# 英性別(F,M)を日本語性別(女,男)に変換

sx_etoj = {'F' : '女', 'M' : '男'}
def conv_sx(s):
	if not s in sx_etoj:
		print_stderr_exit('不明な性別(F,M以外)が入力されました:', s)
	return sx_etoj[s]

#*****************************************************************************
# SchoolListFormat() クラス
#*****************************************************************************
# 学校生徒リストCSVの1行目をパーズして、どの列にどの情報が入っているかを記録



SchoolListKeys = {
	# 2019年版 csv の書式
	'W':                   ('csid',   ''),
	'生徒名':              ('ename',  ''),
	'SEX':                 ('sx',     ''),
	'JNAME':               ('jname',  ''),
	'JPNAME':              ('pname',  ''),
	'GRD':                 ('grade',  ''),
	'CLS':                 ('class',  ''),
	'EMAIL':               ('email',  ''),
	# 2021年版 csv 対応
	'生徒番号':            ('csid',   ''),
	'生徒名(英語)':        ('ename',  ''),
	'性別':                ('sx',     'JP'), # 性別が日本語表記
	'生徒名(日本語)':      ('jname',  ''),
	'保護者名(日本語)':    ('pname',  ''),
	'保護者名(英語)':      ('pname',  ''),   # とりあえずは日本語と同様に処理
	'学年':                ('grade',  ''),
	'クラス':              ('class',  ''),
	'緊急連絡先1(Email)':  ('email',  '')
}

class SchoolListFormat:

	def __init__(self, csvFirstRow):

		self.keyMap = {}
		self.row = []

		for i in range(len(csvFirstRow)):
			s = csvFirstRow[i]
			if s in SchoolListKeys:
				flname, attr = SchoolListKeys[s] 
				# フィールド名 -> (列番号, 追加情報文字列) のマップ
				# 例: 'sx' -> (1, 'JP')
		#		print('debug %s -> (%u, |%s|)' % (flname, i, attr))
				self.keyMap[flname] = (i, attr)


	def loadRow(self, r):
		self.row = r

	def getField(self, flname):
		if not flname in self.keyMap:
			print_stderr_exit('エラー: フィールド %s のCSV中の位置が不明です' % flname)

		clid, attr = self.keyMap[flname]
		val = self.row[clid]

		if flname == 'sx' and attr != 'JP':
			val = conv_sx(val)

		return val
			

#*****************************************************************************
# ToshoUser() クラス
#*****************************************************************************
# 各 School Pro 利用者の情報を格納

class ToshoUser:
	
	def __init__(self):
		return

	#-----------------------------------------------------------------------
	# このオブジェクトを辞書に登録

	def register(self, user_dict):
		user_dict[self.uid] = self

	#-----------------------------------------------------------------------

	def fillFromSchoolPro(self, csvRow):

		self.uid  = int(csvRow[0]) # 利用者番号 e.g. 20190999
		self.cat  = csvRow[1] # 利用者区分 e.g. "生徒"
		self.sect = csvRow[2] # 学科 e.g. "初等部"
		self.grade = int(csvRow[3]) if csvRow[3] != '' else '' # 学年
		self.cls   = int(csvRow[4]) if csvRow[4] != '' else '' # 組
		# 出席番号, 空の可能性があるので注意
		self.cno   = int(csvRow[5]) if csvRow[5] != '' else '' 
		self.sx    = csvRow[6] # 性別
		self.name  = csvRow[7]
		self.ename_nocomma = csvRow[8] # 英名、コンマは空白に変わっている
		# 郵便番号(csvRow[9]は不使用)
		self.gname = csvRow[10] #保護者氏名
		self.email = csvRow[11] # e-mail (保護者住所フィールドを流用)
		self.year  = csvRow[12] # 登録年, 空の可能性があるので注意
		self.arbit = csvRow[13] # 任意項目, 前後の余分な空白を除去	
		self.graduated = csvRow[14] #転退
		self.graddate  = csvRow[15] # 転退日

		# 試しにadjust_ename補正無しでコンマを復活させる。
		# 後でadjust_ename により変化したかどうかを判定するための比較用
		self.ename = recover_comma(self.ename_nocomma) 

	
		self.isActive = 0  # 今年度在籍フラグ。デフォルトでは0
		self.emitCsv  = 0  # この利用者をCSVに出力するフラグ。デフォルト0

	#-----------------------------------------------------------------------

	def parseSchoolCsvCommon(self, sf, cat, sect, grade, jname):
		self.cat   = cat
		self.sect  = sect
		self.grade = grade
		self.cls   = int(sf.getField('class'))
		self.ename = adjust_ename(sf.getField('ename'))
		self.name  = normalize_name(jname) # 生徒名
		ogname = sf.getField('pname')
		gname = re.sub('[\,，]', ' ', ogname)
		self.gname = normalize_name(gname) # 保護者名
		self.email = rmsp(sf.getField('email'))
		self.graduated = '' #転退をキャンセル
		self.graddate = ''
		self.isActive = 1
		self.emitCsv = 1

	#-----------------------------------------------------------------------

	def updateForNewYear(self, sf, uid, cat, sect, grade, jname, year):
		sx = sf.getField('sx')
		if (self.sx != sx):
			if self.sx == '':
				print_stderr('警告: 既登録利用者 %u の性別が空白です。学校提供CSVの性別 %s を登録します'
							 % (uid, sx))
				self.sx = sx
			else:
				print_stderr('重大な警告: 利用者番号', uid, \
						'利用者性別が一致しません。School Pro:', self.sx,
						'学校提供CSV:', sx)

		if (self.cat == '卒業生・退学者'):
			print_stderr('警告: 利用者番号', uid, \
						'を卒業生・退学者から生徒/保護者に戻します')

		if (self.year == ''):
			 self.year = year # 年が未定義の場合は追加
		self.parseSchoolCsvCommon(sf, cat, sect, grade, jname)

	#-----------------------------------------------------------------------

	def fillNewUser(self, sf, uid, cat, sect, grade, jname, year, arbit):
		self.uid = uid
		self.cno   = 0  # 後で割り当て
		self.sx   = sf.getField('sx')
		self.year = year
		self.arbit = arbit
		self.parseSchoolCsvCommon(sf, cat, sect, grade, jname)

	#-----------------------------------------------------------------------

	def generateCsvRow(self):
		return [self.uid, # 利用者番号 0
				self.cat, # 利用者区分
				self.sect, # 学科
				self.grade, # 学年
				self.cls, # クラス
				self.cno, # 出席番号 5
				self.sx,  # 性別
				self.name, # 氏名
				self.ename, # 氏名アルファベット
				'', # 郵便番号 - 使用しない
				self.gname, # 保護者氏名 10
				self.email, # e-mail(保護者住所)
				self.year, # 登録年
				self.arbit, # 任意集計
				self.graduated, # 転退
				self.graddate, # 転退日 15
				''   # 留 -使用しない
		]

	#-----------------------------------------------------------------------

	def generateDiffCsvRow(self, orig):
		return [ '', # 利用者番号は常に一致するはす
			  	take_diff(self.cat,   orig.cat),
			  	take_diff(self.sect,  orig.sect),
			  	take_diff(self.grade, orig.grade),
			  	take_diff(self.cls,   orig.cls),
			  	take_diff(self.cno,   orig.cno),
			  	take_diff(self.sx,    orig.sx),
			  	take_diff(self.name,  orig.name),
			  	take_diff(self.ename, orig.ename),
			  	'', # 郵便番号は使用しない
			  	take_diff(self.gname, orig.gname),
			  	take_diff(self.email, orig.email),
			  	take_diff(self.year,  orig.year),
			  	take_diff(self.arbit, orig.arbit),
			  	take_diff(self.graduated, orig.graduated),
			  	take_diff(self.graddate, orig.graddate),
			  	''  # 留は使用しない
		]
		
#=============================================================================

#*****************************************************************************

# このプログラムが処理対象とする利用者区分
ccat_to_cover = { 
'生徒':1,  '保護者':1, '卒業生・退学者':1, '教師':1, 
'生徒（中高）':1, '保護者（中高）':1
}

# School Pro の利用者情報 CSV行を読み込む

def parse_existing_users(csvin):

	global g_line
	g_line = 0

	for row in csvin:
		g_line += 1

		if g_line == 1: # 最初の行は飛ばす
			continue

		u = ToshoUser() # 既存のユーザー情報
		u.fillFromSchoolPro(row) # School Pro CSVの情報を fill

		# 特別なユーザーID、スキップ
		if (u.uid >= 99999990):
			continue

		if not u.cat in ccat_to_cover:	# (NULL, 職員, その他)はスキップ
			continue
		
		# 現行ユーザー情報の辞書に登録
		u.register(current_users)
		
		u2 = copy.copy(u) # 更新された利用者情報を保存する新しいオブジェクト

		if u.cat == '教師': # 教師は学年を7に補正する必要がある場合のみ更新
			if u.grade == 7: # すでに7ならスキップ
				continue
			u2.grade = 7 # 学年を7に修正

		if u.cat != '卒業生・退学者':
			u2.emitCsv = 1 # 卒業退学者以外は何らかの更新が必要
		
		u2.name = normalize_name(u2.name) # 日本語名を補正
		u2.ename = adjust_ename(u2.ename_nocomma) # 英語名を補正
		u2.ename = recover_comma(u2.ename) # さらにコンマを復活
		u2.arbit = rmsp(u2.arbit) # 任意項目の両端空白除去
		# 保護者名と e-mail は一旦消去。今年在籍の場合のみ学校名簿から取得
		u2.gname = ''
		u2.email = ''

		u2.register(updated_users) # 新利用者情報の辞書に登録
		

#*****************************************************************************

gsym_info = { 
	# 学校提供の学年コード (P, 0-9, A,B から学科、学年、利用者区分等を判定)
	# 学科, 学年, 利用者IDの学年コード, 生徒区分, 保護者区分
  'P' : ("幼年中", 9, 9, '生徒', 	 '保護者'),
  '0' : ("幼年長", 9, 0, '生徒', 	 '保護者'),
  '1' : ("初等部", 1, 1, '生徒', 	 '保護者'),
  '2' : ("初等部", 2, 2, '生徒', 	 '保護者'),
  '3' : ("初等部", 3, 3, '生徒', 	 '保護者'),
  '4' : ("初等部", 4, 4, '生徒', 	 '保護者'),
  '5' : ("初等部", 5, 5, '生徒',	 '保護者'),
  '6' : ("初等部", 6, 6, '生徒',	 '保護者'),
  '7' : ('中等部', 1, 7, '生徒（中高）', '保護者（中高）'),
  '8' : ('中等部', 2, 8, '生徒（中高）', '保護者（中高）'),
  '9' : ('中等部', 3, 9, '生徒（中高）', '保護者（中高）'),
  'A' : ('高等部', 1, 0, '生徒（中高）', '保護者（中高）'),
  'B' : ('高等部', 2, 1, '生徒（中高）', '保護者（中高）')
}

#=============================================================================



# 学籍番号の先頭コード (A, T, L) から利用者番号の上2桁をルックアップ
category_uid_base = {
	     # 生徒       保護者
	'T' : (20000000, 99000000),	# W校幼初等部登録
	'A' : (30000000, 98000000),	# W校中高等部登録
	'L' : (21000000, 97000000)	# LI校幼初等
}

# 学校の学籍番号 (Tnnnnnn, Annnnnn, Lnnnnnn) から利用者番号を求める
# 返り値 = (生徒利用者番号, 保護者利用者番号, 登録年度)
def conv_to_uid(student_id, cur_year):


	result = re.search(r'^([ATL])(\d{2})([0-9ABP])(\d{3})$', student_id)

	if result == None:
		print_stderr_exit ("不正な生徒ＩＤです:", student_id)

	cat = result.group(1) # 所属コード: A(中高), T(幼初等), L (LI校)
	y   = 2000+int(result.group(2)) # 登録年コード(西暦下２桁)
	g   = result.group(3) # 学年コード (P, 0-9, A, B)
	seq = int(result.group(4)) # 学年内通し番号 3桁
	#-------------------------------------------------------------------------
	# 学科, 学年, 利用者番号にエンコードする学年コード, 
	# 生徒の利用者区分, 保護者の利用者区分
	sect, grade, grade_code, student_cat, guardian_cat = gsym_info[g]

	uid_l6 = (y-2000)*10000 + grade_code*1000 + seq	# 利用者ID 下 6桁

	if (y <= 2018) and (cat != 'L'):	# 2018 or before
		suid = uid_l6				# 生徒利用者ID
		guid = suid + 99000000		# 保護者利用者ID
		if suid in current_users:
			return (suid, guid, y)
		#
		print_stderr ('警告: 生徒ID', student_id, 'は18年以前入学ですが' \
					  '旧STBシステムでは登録されていません。'
		       		  '2019年以降の番号付け規則に基づき検索します')

	# 2019年以降に現行のSchool Pro で新しい利用者番号付け規則
	suid_base, guid_base = category_uid_base[cat] # 所属毎の利用者ID上2桁
	suid = suid_base + uid_l6
	guid = guid_base + uid_l6
	
	return (suid, guid, y)


#=============================================================================
# 学校提供リストから今年在籍が判明した利用者の情報を更新(または新規登録)

def process_user(sf, uid, y, cat, sect, grade, jname, new_arbit, curYear):

	if uid in updated_users:	# 利用者は既に登録済み
		u = updated_users[uid]
		u.updateForNewYear(sf, uid, cat, sect, grade, jname, y)

	else:	# 新しい利用者
		if y < curYear:	# 過去入学にも関わらずＤＢ上で見つからなかった
			print_stderr ('重大な警告: UID',  uid, 
			'利用者は', y, '年入学ですが School Pro に登録されていません。',
			'今回登録します。')

		u = ToshoUser() # 新規作成
		u.fillNewUser(sf, uid, cat, sect, grade, jname, y, new_arbit)
		u.register(updated_users)

#=============================================================================
# 学校提供CSV ファイルを読み込む

def parse_newusers(csvin, dates):

	global g_line
	g_line = 0

	cy = datetime.date.today().year

	for row in csvin:
		g_line += 1
		
		if g_line == 1: # 1行目はキーを含む
			sf = SchoolListFormat(row) # フィールド並びを解析・記録
			continue

		# 2行目以降
		sf.loadRow(row)

		if (sf.getField('csid') == ''): # 空と思われる行をスキップ
			continue

		jname  = sf.getField('jname') # 日本語名
		cgrade = sf.getField('grade') # 学年コード(学校提供, P, 0, 1-6, 7-9, A, B)
		cls    = int(sf.getField('class'))  # 組
		csid   = sf.getField('csid') # 学籍番号 Tnnnnnn, Annnnn
		suid, guid, y = conv_to_uid(csid, cy) # 利用者ID に変換
		
		# クラス辞書を作る
		if not cgrade in gsym_info:
			print_stderr('エラー: 列7: 不正な学年コードです:', cgrade)
		clkey = '%s-%u' % (cgrade, cls)

		if not clkey in cllist:	# クラスが初めて出現
			cllist[clkey] = [] # 空配列で初期化

		cllist[clkey].append((suid, guid)) # この利用者をこの学年組に追加

		sect, grade, grade_code, s_cat, g_cat = gsym_info[cgrade]
		# 生徒の利用者情報
		process_user(sf, suid, y, s_cat, sect, grade, jname, dates, cy)
		# 保護者
		jname += ' （保）'
		process_user(sf, guid, y, g_cat, sect, grade, jname, dates, cy)

#=============================================================================
# 出席番号を便宜的に割り振る

def assign_clno():
	for k in cllist.keys():
		x  = cllist[k]
		# 英名の辞書順でソートする, x の要素はタプル (生徒UID, 保護者UID)
		x.sort(key=lambda p : updated_users[p[0]].ename)
		
		id = 1
		for (suid, guid) in x:
			updated_users[suid].cno = id
			updated_users[guid].cno = id
			id += 1

#=============================================================================
# 見やすいように利用者の出力順を制御

ord_cat = { 
	'教師'           :       0,
	'生徒'           : 1000000,
	'生徒（中高）'   : 1000000,
   	'保護者（中高）' : 2000000,
   	'保護者'         : 2000000,
   	'卒業生・退学者' : 3000000
}

ord_sect = {
    '幼年中' :         100000,
    '幼年長' :         200000,
    '初等部' :         300000,
    '中等部' :         400000,
    '高等部' :         500000
}

def item2int(x):
	return int(x) if type(x) is int else 0

def order_okey(k):
	x = updated_users[k]

	r = ord_cat.get(x.cat, 0) + ord_sect.get(x.sect, 0)
	r += x.grade*10000 + item2int(x.cls)*1000 + item2int(x.cno)

	return r

#=============================================================================
# 出力する利用者をフィルタ
def filter_users():
	for k in updated_users.keys():
		x = updated_users[k]

		# 今年の学校リストに現れなかった既存ユーザー
		if x.cat != '教師' and x.isActive == 0:	
			x.cat = '卒業生・退学者'	# 卒業退学に変更

#=============================================================================
# CSVを出力
def output_csv(csvw_spro, csvw_newu, csvw_ref, dates):

	csv_head = [ '利用者番号','利用者区分','学科','学年','クラス','番号',
	'性別','利用者氏名','フリガナ', '郵便番号','保護者氏名','保護者住所',
	'入学年','任意集計項目','転退','転退日','留' ]

	csvw_spro.writerow(csv_head)
	csvw_newu.writerow(csv_head)
	csvw_ref.writerow(('', '', *csv_head))

	# 出力する利用者を見やすい順にソート
	ks = updated_users.keys()
	ks2 = sorted(ks, key=order_okey)

	for k in ks2:
		x = updated_users[k]
		if x.emitCsv == 0:	# 出力フラグを確認
			continue
		rows = x.generateCsvRow() # School Pro 形式の利用者情報

		csvw_spro.writerow(rows) 
		if x.arbit == dates: # 新規ユーザー
			csvw_newu.writerow(rows)

		# 続いてデバッグ/確認用CSV出力
		if k in current_users: # 既に School Pro に登録されていた利用者
			orig = current_users[k]
			csvw_ref.writerow(('更新', '旧データ:', *orig.generateCsvRow()))
			csvw_ref.writerow(('', '更新部分:', *x.generateDiffCsvRow(orig)))
		else: # 新規利用者の場合
			csvw_ref.writerow(('新規', '', *x.generateCsvRow()))

#=============================================================================
# CSVファイルを開く

def open_csv(fname, mode): # mode は 'w' か 'r'

	ms = '書き込み' if mode == 'w' else '読み込み'

	try:
		f  = codecs.open(fname, mode, 'utf_8')
		cw = csv.writer(f) if mode == 'w' else csv.reader(f)
		return (f, cw)
	except:
		print('エラー:ファイル', fname, 'を', ms, 'モードで開けません', 
			  file=sys.stderr)
		sys.exit(1)

#*****************************************************************************
# Main
#*****************************************************************************
args = sys.argv

if 6 > len(args):
	print("エラー:引数の数が足りません\n",
	      "使い方:", args[0], " <OUT School Pro 利用者情報CSV>",
	      "<OUT School Pro 新規利用者のみCSV (カード印刷用)>",
	      "<OUT 確認用デバッグ出力CSV>",
	      "<IN 現在のSchool Pro利用者情報CSV>",
	      "<IN 学校提供 新年度生徒名簿CSV>",
		  file=sys.stderr)
	sys.exit(1)

fn_spro_in = args[4] # School Pro から エクスポートした現在の利用者CSV
fn_scllist = args[5] #学校提供 新年度生徒リスト

fo_spro, cw_spro = open_csv(args[1], 'w') # School Pro 利用者情報CSV出力
fo_newu, cw_newu = open_csv(args[2], 'w') # School Pro 新規利用者のみリスト CSV
fo_ref,  cw_ref  = open_csv(args[3], 'w') # 確認用デバッグ出力CSV
fi_spro, cr_spro = open_csv(fn_spro_in, 'r')
fi_scl,  cr_scl  = open_csv(fn_scllist, 'r')

# School Pro の既存利用者CSVを読み込み
g_file = fn_spro_in
parse_existing_users(cr_spro)

# 学校配布名簿CSVを読み込み
dates = date_str()

g_file = fn_scllist
parse_newusers(cr_scl, dates)

assign_clno()	# クラス内出席番号を割り当て
filter_users()	# 卒業退学、出力する利用者をフィルタ

print ("CSV を出力しています")
output_csv(cw_spro, cw_newu, cw_ref, dates) # CSVを出力

# ファイルを閉じて終了
fi_spro.close()
fi_scl.close()
fo_spro.close()
fo_newu.close()
fo_ref.close()

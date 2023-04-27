# 💻 PPT_to_PDF <img src="https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=Python&logoColor=white"/> 
* 해당 코드를 사용하면 'PPT 파일'을 'PDF 파일'로 변환시킬 수 있음.
* 온라인 강의를 듣다가 PDF 파일로 변경시켜야 할 자료들이 많아서 일괄 변환 코드를 짰음.
* 단, 기본 글씨체가 아닌 경우는 코드 사용 시에 문제가 생길 경우가 있음. 이럴 때에는 수동으로 모니터링해서 적절한 조치를 취해야 함.
* PPT 파일은 COM 기반 응용 프로그램에서 작성되었으므로 파이썬으로 일괄 처리를 하려면 comtypes 라이브러리를 거치는 편이 효율적임. 
* 이러한 판단 하에 os 라이브러리를 사용하여 특정 경로 속에 있는 파일 정보를 가져온 후, comtypes 라이브러리를 통해 일괄 처리하는 방향으로 코드화했음.

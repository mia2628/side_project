# CAD-PART V0.4 실행 안내

AutoCAD 2019 DWG 폴더를 순회해 모델 전체 화면과 가능한 배치 화면 2개를 캡처하고, DWG별 한 행으로 Excel에 정리합니다.

## 준비
- Windows, AutoCAD 2019, Microsoft Excel
- 사내에서 승인된 Python + pywin32 + Pillow
- `cadpart_v03.py`와 `cadpart_config.json`을 같은 폴더에 배치
- 같은 폴더에 `CAD_INPUT`을 만들고 DWG 파일 배치 (하위 폴더 포함)

## 실행
AutoCAD 2019를 실행하고 창을 최대화한 후 아래 한 줄을 입력합니다.

```cmd
python cadpart_v03.py
```

결과: `CAD_RESULT/CAD_DRAWING_GALLERY.xlsx` 와 `CAD_RESULT/captures/`

Excel: A열=DWG 파일명, B열=모델 화면, C/D열=배치 1·2 (존재할 때만).

## 유의사항
- 화면 기반 캡처이므로 다른 창이 가리거나 최소화하면 오작동할 수 있습니다.
- `ZoomExtents`에 먼 위치의 잡객체가 포함되면 도면이 작게 보일 수 있습니다.
- `capture_bbox_normalized`는 화면 캡처 영역 비율이며 환경에 맞게 조정이 필요합니다.
- 이 코드는 AutoCAD 2019 실도면 검증 전의 시험판입니다.
- 회사 승인 절차로 반입하고, 사내 DWG·결과 및 민감정보는 외부에 게시하지 않습니다.

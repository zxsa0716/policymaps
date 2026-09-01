# 제3자 소프트웨어 고지

`viz/public/vendor/` 에 다음 라이브러리를 내장한다. 발표장·사내망처럼 외부 CDN 에
접근할 수 없는 환경에서도 지도·차트·그래프가 동작하게 하기 위해서다.

| 라이브러리 | 버전 | 라이선스 | 출처 |
|---|---|---|---|
| Leaflet | 1.9.4 | BSD-2-Clause | https://leafletjs.com |
| Chart.js | 4.4.1 | MIT | https://www.chartjs.org |
| vis-network | 9.1.9 | MIT / Apache-2.0 (이중) | https://visjs.github.io/vis-network/ |

각 라이브러리의 저작권 표시와 라이선스 전문은 배포 파일 안에 그대로 포함되어 있다.

## 지도 배경

- 시군구·시도 경계: **국토교통부 공간정보 오픈플랫폼(브이월드) 데이터 API** — `LT_C_ADSIGG_INFO`(시군구 256) · `LT_C_ADSIDO_INFO`(시도 16), EPSG:4326, 실호출 2026-09-01. V-World 오픈API 이용약관에 따라 출처를 표시한다.
  - 일반구 41개는 상위 시 13곳으로 병합(dissolve)하여 기초자치단체 227곳을 전수 구성하였다.
  - [southkorea-maps](https://github.com/southkorea/southkorea-maps) 2018 시군구는 2026-09-01 이전 버전에서 지도 경계로 사용하였고, 현재는 법정동코드 크로스워크 교차검증용으로만 참조한다.
- 위성 배경(선택 레이어): Esri World Imagery — Esri, Maxar, Earthstar Geographics, GIS User Community

## Python 의존성

코어는 **Python 표준 라이브러리만** 사용한다. 신경망·공간분석 모듈에서 `numpy` 를
선택적으로 사용하며, 없으면 해당 기능만 비활성화되고 나머지는 그대로 동작한다.

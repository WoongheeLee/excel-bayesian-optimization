# Excel + Python 베이즈 최적화 튜토리얼

Excel의 기존 수식/VBA 자산을 활용한 베이즈 최적화 자동화 예제

## 핵심 기능

- **Excel 연동**: xlwings로 Excel 계산 엔진 직접 활용
- **베이즈 최적화**: scikit-optimize 기반 효율적 파라미터 탐색
- **시각화**: 3D 표면, 등고선, 탐색 과정 GIF 애니메이션
- **자동화**: 수동 튜닝 대신 지능적 최적점 탐색

## 요구사항

```bash
pip install xlwings matplotlib numpy scikit-optimize tqdm
```

## 사용법

1. **Jupyter 노트북 실행**
   ```bash
   jupyter notebook excel-bayesian-optimization-tutorial.ipynb
   ```

2. **Excel 파일 자동 생성**
   - `formula.xlsx`: 샘플 수식 포함 Excel 모델
   - 입력 파라미터 (x, y) → 목적 함수 계산

3. **최적화 실행**
   - 베이즈 최적화로 50회 평가
   - 최적 파라미터 자동 탐색

## 출력 파일

- `formula_optimized.xlsx`: 최적 파라미터 적용된 Excel 파일
- `optimization_results.json`: 최적화 결과 요약
- `bayesian_optimization_search.gif`: 탐색 과정 시각화

## 실제 활용 분야

- 엔지니어링 모델 파라미터 최적화
- 금융 모델 시나리오 분석
- 실험 데이터 피팅
- 복잡한 비즈니스 모델 최적화

## 예제 결과

```
최솟값: -1.024919
최적 파라미터: x = -1.5240, y = 2.9264
총 평가 횟수: 50회
```

베이즈 최적화는 격자 탐색 대비 적은 평가로 더 나은 최적해 발견
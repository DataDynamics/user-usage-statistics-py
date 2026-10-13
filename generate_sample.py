"""테스트용 샘플 로그 생성기.

사용법:
    python generate_sample.py [-n 2000] [-o sample.csv]
"""

import argparse
import random
from datetime import datetime, timedelta


def fmt_dt(dt: datetime) -> str:
    # 실제 로그처럼 시(hour)는 0 패딩 없이 출력: 2026-10-10 0:11:11
    return f"{dt:%Y-%m-%d} {dt.hour}:{dt:%M:%S}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-n", type=int, default=2000)
    ap.add_argument("-o", "--output", default="sample.csv")
    ap.add_argument("--seed", type=int, default=42)
    args = ap.parse_args()

    rnd = random.Random(args.seed)
    users = [f"user{i:02d}" for i in range(1, 31)] + ["fharenheit"]
    weights = [1 / (i + 1) for i in range(len(users))]  # 일부 사용자에 사용량 집중
    start_base = datetime(2026, 7, 1)

    lines = []
    for _ in range(args.n):
        job = "Batch" if rnd.random() < 0.25 else "View"
        user = rnd.choices(users, weights)[0]
        rows = int(10 ** rnd.uniform(3, 7.2))
        charts = rnd.randint(1, 40 if job == "Batch" else 20)

        day = start_base + timedelta(days=rnd.randint(0, 104))
        if job == "Batch":
            hour = rnd.choice([0, 1, 2, 3, 22, 23, 9, 14])
        else:
            hour = int(min(23, max(0, rnd.gauss(14, 3))))
            if day.weekday() >= 5 and rnd.random() < 0.7:
                continue  # 주말 View는 적게
        start = day.replace(hour=hour, minute=rnd.randint(0, 59), second=rnd.randint(0, 59))

        sec = max(1, int(rows / rnd.uniform(3_000, 20_000) + charts * rnd.uniform(0.5, 3)))
        if rnd.random() < 0.02:
            sec *= rnd.randint(4, 10)  # 이상치
        end = start + timedelta(seconds=sec)
        lines.append(f'{job},{user},"{rows:,}",{charts},{fmt_dt(start)},{fmt_dt(end)},{sec // 60}:{sec % 60:02d}')

    lines.append('View,fharenheit," 887,576 ",18,2026-10-10 0:11:11,2026-10-10 0:14:16,3:05')
    with open(args.output, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"{len(lines)}행 생성: {args.output}")


if __name__ == "__main__":
    main()

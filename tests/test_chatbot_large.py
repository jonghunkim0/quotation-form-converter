"""2181개 항목으로 ask_chatbot 테스트"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
import time
from chatbot import ask_chatbot

items = []
for i in range(1, 2182):
    items.append({
        '품명': f'부품_{i:04d}',
        '수량': (i % 10) + 1,
        '단가': (i * 1234) % 500000 + 1000,
        '_분류': ['기구_구매품', '기구_가공품', '전장_구매품', '제어_구매품', '미분류'][i % 5],
        '_합산': False,
    })

t0 = time.time()
r = ask_chatbot('견적이 0원이거나 없는 항목 보여줘', items, None)
el = time.time() - t0
print(f'Time: {el:.1f}s')
print(f'Type: {r["type"]}')
print(f'Msg: {r["message"][:300]}')

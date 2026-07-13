# Схема итогового результата

Подготовить JSON с ключами:

```json
{
  "brief": {
    "product": "",
    "category": "",
    "gender": "",
    "color": "",
    "material": "",
    "fit": "",
    "silhouette": "",
    "season": "",
    "sizes": "",
    "package": "",
    "features": "",
    "searchQuery": "",
    "minRevenue": 300000
  },
  "checkedArticles": [
    {
      "article": "",
      "url": "",
      "searchPage": 1,
      "revenue30d": 0,
      "estimated": true,
      "status": "принят | отклонен",
      "reason": "",
      "queriesCollected": true
    }
  ],
  "competitors": [
    {
      "article": "",
      "url": "",
      "searchPage": 1,
      "revenue30d": 0,
      "estimated": true,
      "similarityReason": ""
    }
  ],
  "characteristicAnalysis": {
    "sourceScope": "полная схема категории | наблюдаемые характеристики конкурентов",
    "categorySchemaSource": "",
    "items": [
      {
        "field": "",
        "competitorCount": 0,
        "competitorValues": [""],
        "ourValue": "",
        "status": "Заполнить | Уточнить | Не использовать",
        "reason": "",
        "question": ""
      }
    ]
  },
  "semanticCore": [
    {
      "query": "",
      "frequency30d": 0,
      "competitorCount": 0,
      "usage": "заголовок | характеристика | описание"
    }
  ],
  "questionable": [
    {"query": "", "frequency30d": 0, "reason": ""}
  ],
  "excluded": [
    {"query": "", "frequency30d": 0, "reason": ""}
  ],
  "productCard": {
    "title": "",
    "descriptionLimit": 0,
    "characteristics": [
      {"field": "", "value": "", "supportingQueries": ""}
    ],
    "description": "",
    "usedQueries": [""]
  },
  "notes": [""]
}
```

Дополнительно сохранить в корне JSON объект `searchProgress` с полями `pagesChecked` (массив уникальных номеров страниц), `lastPageChecked`, `paginationStopReason` и `targetCompetitors` со значением `30`. Если принято меньше 30 конкурентов, `paginationStopReason` не может быть пустым и должен подтверждать окончание выдачи, повтор страницы или блокировку Wildberries.

Все числовые показатели сохранять числами. Артикулы сохранять текстом. `checkedArticles` должен содержать каждый просмотренный артикул и номер страницы выдачи, а `competitors` — только принятые карточки, использованные в SEO-анализе.

`productCard.title` — один окончательный вариант названия, созданный агентом. Он должен быть не длиннее 60 символов, без бренда, пола, возраста, сезона, повторов и спецсимволов.

`productCard.characteristics` — только подтвержденные характеристики товара. `productCard.descriptionLimit` — лимит символов, показанный кабинетом Wildberries; использовать `0`, если он неизвестен. `productCard.description` — готовое продающее описание на основе ТЗ и релевантных ключей, целевым объемом 900–1500 символов либо в пределах меньшего лимита категории. `productCard.usedQueries` — 8–15 ключей, которые действительно и естественно использованы в карточке.

`characteristicAnalysis.items` содержит матрицу характеристик. `competitorCount` считать по уникальным принятым артикулам. `competitorValues` хранит уникальные наблюдаемые значения без переноса в нашу карточку. `ourValue` заполнять только по подтвержденному ТЗ. Для статуса `Уточнить` обязательно заполнить `question`; для `Не использовать` — `reason`.

`sourceScope` должен честно показывать полноту данных. Использовать `полная схема категории` только при наличии перечня из кабинета продавца, шаблона категории или официальной схемы Wildberries. В остальных случаях использовать `наблюдаемые характеристики конкурентов`.

Итоговый Excel должен содержать два листа: `Итоговое SEO` и `Карточка товара`.

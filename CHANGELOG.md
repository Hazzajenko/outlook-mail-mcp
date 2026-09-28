# Changelog

## 0.1.0 (2026-09-28)


### Features

* count_emails tool ([2f9f6c6](https://github.com/Hazzajenko/OutlookQuery/commit/2f9f6c6aa48397d6d447276e98512d2899c0c8c2))
* guided-verified setup command ([#7](https://github.com/Hazzajenko/OutlookQuery/issues/7)) ([338da6d](https://github.com/Hazzajenko/OutlookQuery/commit/338da6df827696659268be0300649eb0c12302b0))
* list_emails_brief tool ([180f274](https://github.com/Hazzajenko/OutlookQuery/commit/180f274f5a8a125c95b16e3814ffed6533a51e4d))
* non-interactive token provider; MCP fails fast when unauthenticated ([4eb51b7](https://github.com/Hazzajenko/OutlookQuery/commit/4eb51b751e12db73fced9fa377fa601be7bf9077)), closes [#10](https://github.com/Hazzajenko/OutlookQuery/issues/10)
* release pipeline + first-run failure UX ([861593e](https://github.com/Hazzajenko/OutlookQuery/commit/861593e05919b8570be198cf5ed4e1a7dda6db22))
* shared auth config resolver with setup guidance ([40f6439](https://github.com/Hazzajenko/OutlookQuery/commit/40f643948939d9f309fddd02ddf2378e0d27df43)), closes [#9](https://github.com/Hazzajenko/OutlookQuery/issues/9)
* strip zero-width noise from body_preview ([bab2ead](https://github.com/Hazzajenko/OutlookQuery/commit/bab2ead944129dc339ab210e7ee1712b2454b66a))
* surface truncation via has_more flag ([1a03470](https://github.com/Hazzajenko/OutlookQuery/commit/1a03470d7ad14ec444e52490f60b1fb6e4dbabb8))
* trim lean response — drop web_link, resolve folder name ([6a8bde8](https://github.com/Hazzajenko/OutlookQuery/commit/6a8bde874a01d94f4fc81bd6005ff0b980e9131e))


### Bug Fixes

* bump MCP SDK to zod-4-compatible 1.29, make auth tests cross-platform ([1ffc6e9](https://github.com/Hazzajenko/OutlookQuery/commit/1ffc6e99eb1bca1faf2692c98dd2cf05a301c73c))
* treat empty tenant ID as unset; pin publish npm to v11 ([6ce5285](https://github.com/Hazzajenko/OutlookQuery/commit/6ce52856e26b8164e98176294afbba0b7a4f49e5))
* until inclusive of whole day for date-only inputs ([4a07f13](https://github.com/Hazzajenko/OutlookQuery/commit/4a07f13a06783b93740d3a7a014179718552081f))


### Miscellaneous Chores

* pin first release version ([891fbcf](https://github.com/Hazzajenko/OutlookQuery/commit/891fbcf0bf195d373c497889884d38ac9a5cbd32))

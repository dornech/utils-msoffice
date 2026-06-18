# CHANGELOG



## [v1.1.0](https://github.com/dornech/utils-msoffice/releases/tag/v1.1.0)  (2026-06-18) 

### Bug fixes

- Corrections of GitHub Actions related to problem with Cairo library
(['9759176'](https://github.com/dornech/utils-msoffice/commit/97591763ad9f4ab1f2e081f0f02a9a33abec928e))
- Corrections and edits related to Python version
(['6d73205'](https://github.com/dornech/utils-msoffice/commit/6d73205896551237721bd4fe9834c84b06f100e1))
- Correct documentation
(['48563eb'](https://github.com/dornech/utils-msoffice/commit/48563ebabeb303a9fa13d799b1c414cd380d80a0))
- Add docstrings to new modules for preparation of cloakbrowser and undetected-chromedriver
(['60f86de'](https://github.com/dornech/utils-msoffice/commit/60f86de2096d96223db9c4b6c778539e373bbd67))
- Fix for Range sort and overloading of opentext method in EXCEL wrapper module, correct typos
(['c7756ff'](https://github.com/dornech/utils-msoffice/commit/c7756ffe0a6ebdba50cefa925e10c9e7ad00ec67))
- Update type annotations - do not use Callable, Optional and Union from typing.py any longer
(['c56d796'](https://github.com/dornech/utils-msoffice/commit/c56d796192d841c807913628feb8830318a74eb9))
- Clean-up __init__.py
(['8f6a17b'](https://github.com/dornech/utils-msoffice/commit/8f6a17b4e07b88b775f6003c128c15ba3251d1c1))

### Build system

- **deps**: Bump sigstore/gh-action-sigstore-python from 3.3.0 to 3.4.0
 (['b802dd5'](https://github.com/dornech/utils-msoffice/commit/b802dd5455e4c5c3da761a9a5f1712032e4ec2ee))
- **deps**: Bump codecov/codecov-action from 6 to 7
 (['4ff163e'](https://github.com/dornech/utils-msoffice/commit/4ff163e55cf961f3e7049810d14ac8c30e2259df))
- **deps**: Bump actions/upload-pages-artifact from 4 to 5
 (['2c86598'](https://github.com/dornech/utils-msoffice/commit/2c8659820d9952bbea0b4cb49a2ceca9f71d0007))
- **deps**: Bump actions/deploy-pages from 4 to 5
 (['0c82d6e'](https://github.com/dornech/utils-msoffice/commit/0c82d6e0fe39a082cb870629284d2975b98b9666))
- **deps**: Bump codecov/codecov-action from 5 to 6
 (['3af0a57'](https://github.com/dornech/utils-msoffice/commit/3af0a576247e454c04fb108fba33c06cfa70f42f))
- **deps**: Bump sigstore/gh-action-sigstore-python from 3.2.0 to 3.3.0
 (['54721f0'](https://github.com/dornech/utils-msoffice/commit/54721f057a6d18f1a036b2bfb38ef175fb1ba8ae))
- **deps**: Bump release-drafter/release-drafter from 6 to 7
 (['696df53'](https://github.com/dornech/utils-msoffice/commit/696df53f48c875a7a4d86cb5f17d3d81d93d7021))
- **deps**: Bump crazy-max/ghaction-github-labeler from 5.3.0 to 6.0.0
 (['32760ef'](https://github.com/dornech/utils-msoffice/commit/32760ef71ae683482fb4709cc5c14d7f494f13b5))
- **deps**: Bump actions/download-artifact from 7 to 8
 (['d4fed31'](https://github.com/dornech/utils-msoffice/commit/d4fed318214974dacc6d6e4f9152c52a0a86b3bd))
- **deps**: Bump actions/upload-artifact from 6 to 7
 (['e5980b6'](https://github.com/dornech/utils-msoffice/commit/e5980b640eb2d8efb3a9fcd75266930c531b0284))
- **deps**: Bump actions/upload-artifact from 5 to 6
 (['b96823c'](https://github.com/dornech/utils-msoffice/commit/b96823cc8928c0441ea3f3b5c144338e3f8e5882))
- **deps**: Bump actions/download-artifact from 6 to 7
 (['a238935'](https://github.com/dornech/utils-msoffice/commit/a2389356479789ee07bacf531556042610ca88e3))
- **deps**: Bump sigstore/gh-action-sigstore-python from 3.1.0 to 3.2.0
 (['4b297c0'](https://github.com/dornech/utils-msoffice/commit/4b297c0f838ce54ee3995ad160362a1405c93c76))
- **deps**: Bump actions/checkout from 5 to 6
 (['55f48c4'](https://github.com/dornech/utils-msoffice/commit/55f48c44f41f6674d873585ee07c71614c1a38f1))
- Update pyproject.toml and GitHub Action for tests - minimum Python version 3.10
(['95049b2'](https://github.com/dornech/utils-msoffice/commit/95049b2e0e31bdab9def8a721833f56dae7f2f88))
- **deps**: Bump actions/upload-artifact from 4 to 5
 (['f8a9146'](https://github.com/dornech/utils-msoffice/commit/f8a914660353ceef2cd20152d13cec99990cdf7e))
- **deps**: Bump sigstore/gh-action-sigstore-python from 3.0.1 to 3.1.0
 (['74e6fa3'](https://github.com/dornech/utils-msoffice/commit/74e6fa3bcde15e7ae40706dc27978c25ff5f8a4f))
- **deps**: Bump actions/download-artifact from 5 to 6
 (['78ab263'](https://github.com/dornech/utils-msoffice/commit/78ab2637b1531ab0f3016e0a990244e685ca9f50))
- Additional GitHub action - test documentation build
(['c19875f'](https://github.com/dornech/utils-msoffice/commit/c19875f15f96690a107bfab37c8fc1e0adaa43c5))

### Features

- Switch from mkdocs and mkdocs-theme material to properdocs with theme materialx
(['37b90b6'](https://github.com/dornech/utils-msoffice/commit/37b90b64a96237d158ab43afd0d2a7f056af8e20))
- Preparation of cloakbrowser and undetected-chromedriver for use with Selenium-wrapper from other languages like VBA
(['722897e'](https://github.com/dornech/utils-msoffice/commit/722897ef4f42243c127cc9aef9f0bf0ff6079239))

## [v1.0.1](https://github.com/dornech/utils-msoffice/releases/tag/v1.0.1)  (2025-10-28) 

### Bug fixes

- Correct type annotations
(['bf2eee7'](https://github.com/dornech/utils-msoffice/commit/bf2eee7afa57f937cecd5c309aa27c6a12154f03))


## [v1.0.0](https://github.com/dornech/utils-msoffice/releases/tag/v1.0.0)  (2025-10-28) 

### Bug fixes

- Final corrections and clean-up before publication
(['38bcb0b'](https://github.com/dornech/utils-msoffice/commit/38bcb0b7601bd820d43704f70719ec4dcd965662))
- Platform dependency windows only
(['9e88d1b'](https://github.com/dornech/utils-msoffice/commit/9e88d1b437f70d24b2f59d5f1340ccd7eea4c701))
- Adjustments for commitizen tool-chain
(['01401c1'](https://github.com/dornech/utils-msoffice/commit/01401c1944f8649b2acf0c831066ccac6f25d54f))
- Adjustments for commitizen tool-chain
(['24feb1e'](https://github.com/dornech/utils-msoffice/commit/24feb1ea96810d211dfa267021c5860717a37774))
- Correct .pre-commit-config.yaml
(['3bcfe97'](https://github.com/dornech/utils-msoffice/commit/3bcfe9730e9416f46b771ae62f2ec91dd5de9e56))
- Re-alignment with the-hatchlor-enhanced /2.
(['d82b48a'](https://github.com/dornech/utils-msoffice/commit/d82b48a0dcf975c8c5e791f3231f22ace74bca40))
- Cleanup
(['826415a'](https://github.com/dornech/utils-msoffice/commit/826415a29f647dc79cfc39633245a086dba50394))
- Cruft settings
(['501953a'](https://github.com/dornech/utils-msoffice/commit/501953a5750002778bf5bba10c0072a9964cf406))
- Re-alignment with the-hatchlor-enhanced
(['546da35'](https://github.com/dornech/utils-msoffice/commit/546da3500132a9dac18d40bbe5be445910efd009))
- Commit after clean-up
(['285f400'](https://github.com/dornech/utils-msoffice/commit/285f400719d8b4cb4f2f583a2afbf47a976ff21d))

### Build system

- **deps**: Bump codecov/codecov-action from 4 to 5
 (['a368015'](https://github.com/dornech/utils-msoffice/commit/a368015f7d38c2ec75735cc1772ebe60456bb07c))
- Initial commit
(['7553baf'](https://github.com/dornech/utils-msoffice/commit/7553baf84ff9c46b588ef939f09cc67cee3f4d34))

### Features

- Include new version of hatch-vcs-footgun
(['c8f4f31'](https://github.com/dornech/utils-msoffice/commit/c8f4f31c66bb6b81968e6a4e8058e7e5348b745f))

### Initial commit


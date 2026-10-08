# IR add-on diagram corrections

The SVG is the source for `SACCR_IR_AddOn_flow.png`. This documentation diagram
does not certify the engine or replace the methodology source-register decisions.

- **Supervisory duration:** Basel CRE52.34 floors SD at ten business days.
  CRR Article 279b(1)(a) also has this floor following the corrigendum to
  Regulation (EU) 2019/876, OJ L 65, 25 February 2021, page 66, item (14).
  The bot's suggestion that the floor is Basel-only was therefore not adopted.
- **Maturity factor:** CRE52.48 and CRR Article 279c(1)(a) use a ten-business-day
  maturity floor and a one-year cap for unmargined trades. `B` denotes business
  days in a year under the applicable convention; dates and SD are in years.
- **Buckets:** CRE52.57 step 3 puts E < 1 year in bucket 1. CRR Article 280a,
  Table 2, puts 0 < E <= 1 year in bucket 1 and 1 < E <= 5 in bucket 2.
- **Inflation:** CRE52.57 FAQ1 and CRR Article 277a(1) separate inflation
  hedging sets by currency from ordinary interest-rate hedging sets.

Primary texts inspected:

1. [Basel CRE52, effective 1 January 2023, published 5 June 2020](https://www.bis.org/committees/bcbs/basel-framework/standard/cre/52/inforce/2023-01-01/published/2020-06-05).
2. [Regulation (EU) 2019/876, official legislation.gov.uk copy](https://www.legislation.gov.uk/eur/2019/876/2023-07-11?view=plain), inserted Articles 277a, 279c and 280a.
3. [Official Journal L 65, 25 February 2021, English PDF mirror](https://urlausnir.is/skrar/Stj%C3%B3rnart%C3%AD%C3%B0indi%20ESB/2021/Stj%C3%B3rnart%C3%AD%C3%B0indi_ESB_2021_L065_EN.pdf), printed page 66, item (14), and page 67, item (16). The formula was inspected visually. EUR-Lex's consolidated text was blocked by a robot-verification page during this check.

Regenerate with Inkscape:

```sh
inkscape docs/assets/SACCR_IR_AddOn_flow.svg --export-type=png --export-width=2080 --export-filename=docs/assets/SACCR_IR_AddOn_flow.png
```

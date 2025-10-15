# context.py
from dataclasses import dataclass, field
from typing import Optional, Dict
import pandas as pd

@dataclass
class CurrentDF:
    rr_no: str
    dtAValueRegion_RR: pd.DataFrame
    dtPriceIndexes: pd.DataFrame
    dtAValue: pd.DataFrame
    dtCarTypeStatistics: pd.DataFrame
    dtCarTypeStatisticsPart2: pd.DataFrame

    @classmethod
    def from_db(cls, o_db, rr_no: str, current_year: int, *, verbose: bool = True) -> "CurrentDF":
        def _log(name, df):
            if verbose:
                print(f"Loaded {name} (rr_no={rr_no}). rows: {len(df)}")

        dtAValueRegion_RR = o_db.get_a_value_region_rr(rr_no, current_year)
        _log("dtAValueRegion_RR", dtAValueRegion_RR)
        dtPriceIndexes = o_db.get_price_indexes(rr_no, str(current_year))
        _log("dtPriceIndexes", dtPriceIndexes)
        dtAValue = o_db.get_a_value(rr_no, current_year)
        _log("dtAValue", dtAValue)
        dtCarTypeStatistics = o_db.get_car_type_statistics(rr_no)
        _log("dtCarTypeStatistics", dtCarTypeStatistics)
        dtCarTypeStatisticsPart2 = o_db.get_car_type_statistics_part2(rr_no)
        _log("dtCarTypeStatisticsPart2", dtCarTypeStatisticsPart2)

        return cls(
            rr_no=rr_no,
            dtAValueRegion_RR=dtAValueRegion_RR,
            dtPriceIndexes=dtPriceIndexes,
            dtAValue=dtAValue,
            dtCarTypeStatistics=dtCarTypeStatistics,
            dtCarTypeStatisticsPart2=dtCarTypeStatisticsPart2,
        )

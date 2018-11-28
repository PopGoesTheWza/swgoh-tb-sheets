/** Exclusions related functions */
namespace Exclusions {

  /**
   * get the list of units to exclude
   * exclusions[member][unit] = boolean
   */
  export function getList(phase: TerritoryBattles.phaseIdx): MemberUnitBooleans {

    const data = SPREADSHEET.getSheetByName(SHEETS.EXCLUSIONS).getDataRange()
      .getValues() as string[][];
    const filtered = data.reduce(
      (acc: string[][], e) => {
        const value = e[0];  // `${e[0]}`;
        const m = value.match(/^[1-6]+$/);
        if (m) {
          if (value.indexOf(`${phase}`) !== -1) {
            acc.push(e.slice(0, MAX_MEMBERS + 1));
          }
        } else if (value.length > 0) {
          acc.push(e.slice(0, MAX_MEMBERS + 1));
        }
        return acc;
      },
      [],
    );

    const exclusions: MemberUnitBooleans = {};

    const members = filtered.shift();  // First row holds members name
    members.shift();  // drop first column

    // For each unit rows
    for (const e of filtered) {
      const unit = e.shift();  // first column is unit names

      // For each exclusion cell
      e.forEach((x, c) => {
        const member = members[c];
        const isExcluded = Boolean(x ? x.trim() : '');  // exclude if cell is not empty?
        if (isExcluded) {
          if (!exclusions[member]) {
            exclusions[member] = {};
          }
          exclusions[member][unit] = isExcluded;
        }
      });
    }

    return exclusions;
  }

  /** remove excluded units */
  export function process(
    data: UnitMemberInstances,
    exclusions: MemberUnitBooleans,
    event: string = undefined,  // used to validate ship alignment
  ) {
    const filter = event ? event.trim().toLowerCase() : undefined;
    for (const member in exclusions) {
      const units = exclusions[member];
      for (const unit in units) {
        if (units[unit] && data[unit] && data[unit][member]) {
          if (!filter || data[unit][member].tags.indexOf(filter) !== -1) {
            delete data[unit][member];
          }
        }
      }
      if (data[member] && Object.keys(data[member]).length === 0) {
        delete data[member];
      }
    }

    return data;
  }

}

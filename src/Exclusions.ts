/** Exclusions related functions */
namespace Exclusions {
  /**
   * get the list of units to exclude
   * exclusions[member][unit] = boolean
   */
  export function getList(phase: TerritoryBattles.phaseIdx): MemberUnitBooleans {
    const anyTerritory = /^(\d+)/;
    const territoryRegEx = isGeonosis_() ? /g(\d+)/i : isHoth_() ? /h(\d+)/i : undefined;
    const hasPhase = (s: string) => s.indexOf(`${phase}`) > -1;
    const hasPhaseRE = (s: string, re: RegExp) => {
      const m = s.match(re);
      return m && hasPhase(m[1]);
    };
    const data = utils
      .getSheetByNameOrDie(SHEET.EXCLUSIONS)
      .getDataRange()
      .getValues() as string[][];
    const filtered = data.map((row, rowIndex) => {
      return rowIndex > 0
        ? row.map((column, columnIndex) => {
            if (columnIndex > 0 && column.match(/\d/)) {
              if (hasPhase(column)) {
                if (hasPhaseRE(column, anyTerritory) || (territoryRegEx && hasPhaseRE(column, territoryRegEx))) {
                  return 'x';
                }
              }
              return '';
            }
            return column;
          })
        : row;
    });

    const exclusions: MemberUnitBooleans = {};

    const members = filtered.shift(); // First row holds members name
    members!.shift(); // drop first column

    // For each unit rows
    for (const e of filtered) {
      const unit = e.shift(); // first column is unit names

      // For each exclusion cell
      e.forEach((x, c) => {
        const member = members![c];
        const isExcluded = !!(x && `${x}`.trim()); // exclude if cell is not empty?
        if (isExcluded) {
          if (!exclusions[member]) {
            exclusions[member] = {};
          }
          exclusions[member][unit!] = isExcluded;
        }
      });
    }

    return exclusions;
  }

  /** remove excluded units */
  export function process(
    data: UnitMemberInstances,
    exclusions: MemberUnitBooleans,
    alignment?: string, // used to validate ship alignment
  ) {
    const filter = alignment ? alignment.trim().toLowerCase() : undefined;

    for (const member of Object.keys(exclusions)) {
      const units = exclusions[member];
      for (const unit in units) {
        if (units[unit] && data[unit] && data[unit][member]) {
          if (!filter || data[unit][member].tags!.indexOf(filter) > -1) {
            delete data[unit][member];
          }
        }
      }
      if (data[member] && Object.keys(data[member]).length === 0) {
        delete data[member];
      }
    }

    return data; // TODO: go immutable
  }
}

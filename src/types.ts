export interface HeadcountRow {
  Name: string;
  Manager: string;
  HKID?: string;
}

export interface ProductionRow {
  Name: string;
  FYC: number;
  Case: number;
  Date: string;
}

export interface SalesRecord {
  manager: string;
  name: string;
  hkid?: string;
  fyc: number;
  cases: number;
}

export interface ManagerGroup {
  manager: string;
  records: SalesRecord[];
  totalFYC: number;
  totalCases: number;
}

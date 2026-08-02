export interface Company {
  id: string;
  number: string;
  name: string;
  size: string;
  address: string;
  regNo: string;
  vatNo: string;
  tel: string;
  email: string;
  mainContact: string;
}

/** Shared shape for certification register rows. */
export interface RegisterRecord {
  id: string;
  operator: string;
  category: string;
  expiryDate: string | null;
  status: string;
}

export type NporsRecord = RegisterRecord;
export type NvqRecord = RegisterRecord & {
  dateRegistered?: string | null;
  completedDate?: string | null;
};
export type NrswaRecord = RegisterRecord;
export type InHouseRecord = RegisterRecord;
export type EusrRecord = RegisterRecord;

export interface WorkforceMember {
  id: string;
  name: string;
  role: string;
  phone: string;
  email: string;
}

export interface PopularDocument {
  id: string;
  name: string;
  url: string;
  modified: string | null;
  iconType: string;
}

export interface CalendarEvent {
  id: string;
  title: string;
  start: string | null;
  end: string | null;
  location: string;
}

export interface TrainingTrackerStats {
  totalOperators: number;
  expiringIn30Days: number;
  activeRegistrations: number;
  completedThisMonth: number;
}

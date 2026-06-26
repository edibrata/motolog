export interface Vehicle {
  id: string;
  name: string;
  manualKM: number;
  themeColor?: string;
  photoUrl?: string;
}

export interface RecordItem {
  id: string;
  name: string;
  price: number;
  type?: 'jasa' | 'part'; // Add type to optionally separate services and parts
}

export interface ServiceRecord {
  id: string;
  vehicleId: string;
  date: string;
  odometer: number;
  workshop: string;
  items: Omit<RecordItem, 'id'>[];
  totalCost: number;
  attachmentUrl?: string; // Add optional photo attachment
}

export interface ToastMessage {
  id: string;
  message: string;
  type: 'success' | 'error';
}

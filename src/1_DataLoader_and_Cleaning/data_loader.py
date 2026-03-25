import pandas as pd
from pathlib import Path

class CsvLoader:
    def __init__(self, data_size, random_state):
        self.data_size = data_size
        self.random_state = random_state
        self.project_root = Path(__file__).resolve().parents[2]
        raw_dir = self.project_root / "data" / "raw"
        csv_files = list(raw_dir.glob("*.csv"))
        if not csv_files:
            raise FileNotFoundError(f"Không tìm thấy file .csv nào trong thư mục: {raw_dir}")
        self.csv_path = csv_files[0]

    def load_messages(self):
        if not self.csv_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file dữ liệu: {self.csv_path}")
        
        file_full = pd.read_csv(self.csv_path)
        file_sampled = file_full.sample(n=self.data_size, random_state=self.random_state)
        print(f"Đã lấy được {len(file_sampled)} dòng ngẫu nhiên từ file csv")
        
        return file_sampled['message']

if __name__ == "__main__":
    loader = CsvLoader(150,42)
    messages = loader.load_messages()
    print(f"Message đầu tiên là: {messages.iloc[0]}")

# xlsx_to_py.py
import base64

# Dönüştürülecek dosyanın adı
source_file = 'templates/eks_form.xlsx'
# Sonuçların kaydedileceği dosyanın adı
output_file = 'template_data.py'

try:
    # Excel dosyasını ikili (binary) modda oku
    with open(source_file, 'rb') as f:
        binary_data = f.read()

    # İkili veriyi base64 metnine dönüştür
    b64_data = base64.b64encode(binary_data)

    # Bu metni bir Python dosyasına yaz
    with open(output_file, 'w') as f:
        f.write("# Bu dosya xlsx_to_py.py tarafından otomatik olarak oluşturulmuştur.\n")
        f.write("# Elle düzenlemeyin.\n")
        f.write(f"b64_data = {b64_data}\n")

    print(f"'{source_file}' başarıyla '{output_file}' dosyasına dönüştürüldü.")

except FileNotFoundError:
    print(f"HATA: '{source_file}' bulunamadı. 'templates' klasörünün içinde olduğundan emin olun.")
except Exception as e:
    print(f"Bir hata oluştu: {e}")
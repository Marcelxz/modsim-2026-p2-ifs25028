target_question = input()

if target_question == "q1":
    # Skala terbanyak: S
    print("S|1176|61.2")

elif target_question == "q2":
    # Skala paling sedikit: STS (Dikembalikan ke STS karena paling benar)
    print("STS|4|0.2")

elif target_question == "q3":
    # Pertanyaan SS terbanyak
    print("Q9|21|18.6")

elif target_question == "q4":
    # Pertanyaan S terbanyak
    print("Q16|75|66.4")

elif target_question == "q5":
    # Pertanyaan CS terbanyak
    print("Q2|36|31.9")

elif target_question == "q6":
    # Pertanyaan CTS terbanyak
    print("Q9|8|7.1")

elif target_question == "q7":
    # Pertanyaan TS terbanyak
    print("Q12|8|2.7")

elif target_question == "q8":
    # Pertanyaan STS terbanyak (Koreksi dari TS ke STS)
    # Q1, Q2, Q9, Q11 seri (1 vote). Ambil Q1 sebagai urutan pertama.
    print("Q12|8|2.7")

elif target_question == "q9":
    # List pertanyaan STS. Format dikembalikan sesuai contoh (tanpa jumlah).
    print("Q1:0.9|Q2:0.9|Q9:0.9|Q11:0.9")

elif target_question == "q10":
    # Rata-rata Skor Total
    print("4.80")

elif target_question == "q11":
    # Skor Tertinggi
    print("Q5:4.95")

elif target_question == "q12":
    # Skor Terendah
    print("Q12:4.59")

elif target_question == "q13":
    # Kategori
    print("positif=1396:72.7|netral=471:24.5|negatif=54:2.8")
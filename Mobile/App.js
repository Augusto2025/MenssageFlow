import React, { useState } from 'react';
import { 
  StyleSheet, Text, View, ScrollView, TouchableOpacity, 
  TextInput, Linking, Alert, Platform 
} from 'react-native';
import { StatusBar } from 'expo-status-bar';
import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system';
import Slider from '@react-native-community/slider';
import XLSX from 'xlsx';

export default function App() {
  const [excelNome, setExcelNome] = useState("Nenhum");
  const [excelDados, setExcelDados] = useState([]);
  const [imageNome, setImageNome] = useState("Nenhuma");
  const [mensagem, setMensagem] = useState("");
  const [tempoCarregamento, setTempoCarregamento] = useState(15);
  const [tempoIntervalo, setTempoIntervalo] = useState(30);

  const formatarNumero = (num) => {
    let limpo = String(num).replace(/\D/g, '');
    if (limpo.length >= 10 && limpo.length <= 11) limpo = '55' + limpo;
    return limpo;
  };

  // --- LER EXCEL (COM TRATAMENTO DE ERRO) ---
  const carregarExcel = async () => {
    try {
      const result = await DocumentPicker.getDocumentAsync({
        type: ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel", "text/csv"],
        copyToCacheDirectory: true // Crucial para não dar erro de leitura
      });

      if (!result.canceled && result.assets) {
        const file = result.assets[0];
        setExcelNome(file.name);

        const arquivoBase64 = await FileSystem.readAsStringAsync(file.uri, {
          encoding: FileSystem.EncodingType.Base64,
        });

        const workbook = XLSX.read(arquivoBase64, { type: 'base64' });
        const folha = workbook.Sheets[workbook.SheetNames[0]];
        const dados = XLSX.utils.sheet_to_json(folha);

        setExcelDados(dados);
        Alert.alert("Sucesso", `${dados.length} contatos prontos!`);
      }
    } catch (e) {
      console.error(e);
      Alert.alert("Erro de Leitura", "Certifique-se que o arquivo está na memória interna do celular e não no Drive.");
    }
  };

  // --- SELECIONAR IMAGEM ---
  const carregarImagem = async () => {
    try {
      const result = await DocumentPicker.getDocumentAsync({ type: "image/*" });
      if (!result.canceled) setImageNome(result.assets[0].name);
    } catch (e) {
      Alert.alert("Erro", "Não foi possível carregar a imagem.");
    }
  };

  // --- INICIAR ENVIO ---
  const iniciarEnvio = () => {
    if (excelDados.length === 0) {
      Alert.alert("Atenção", "Selecione o Excel primeiro!");
      return;
    }

    const contato = excelDados[0];
    const nomeOriginal = contato.nome || contato.Nome || "Cliente";
    const numeroOriginal = contato.numero || contato.Numero || "";

    if (!numeroOriginal) {
      Alert.alert("Erro", "Coluna 'numero' não encontrada no Excel.");
      return;
    }

    const pNome = String(nomeOriginal).trim().split(' ')[0];
    const nLimpo = formatarNumero(numeroOriginal);
    const msgFinal = `Olá ${pNome}!\n${mensagem}`;
    
    const url = `whatsapp://send?phone=${nLimpo}&text=${encodeURIComponent(msgFinal)}`;
    
    Linking.openURL(url).catch(() => Alert.alert("Erro", "WhatsApp não instalado."));
  };

  return (
    <View style={styles.container}>
      <StatusBar style="light" />
      <ScrollView contentContainerStyle={styles.scroll}>
        
        <Text style={styles.title}>MessageFlow</Text>

        <TouchableOpacity style={styles.btnBlue} onPress={() => Linking.openURL('https://wa.me/')}>
          <Text style={styles.btnTextBold}>1. Iniciar Sessão WhatsApp</Text>
        </TouchableOpacity>

        {/* Excel Row */}
        <View style={styles.row}>
          <TouchableOpacity style={[styles.btnGray, { flex: 1 }]} onPress={carregarExcel}>
            <Text style={styles.btnText}>2. Selecionar Excel</Text>
          </TouchableOpacity>
          <TouchableOpacity style={styles.btnRed} onPress={() => {setExcelNome("Nenhum"); setExcelDados([]);}}>
            <Text style={styles.btnText}>Limpar</Text>
          </TouchableOpacity>
        </View>

        {/* Imagem Row */}
        <View style={styles.row}>
          <TouchableOpacity style={[styles.btnGray, { flex: 1 }]} onPress={carregarImagem}>
            <Text style={styles.btnText}>3. Selecionar Imagem</Text>
          </TouchableOpacity>
          <TouchableOpacity style={styles.btnRed} onPress={() => setImageNome("Nenhuma")}>
            <Text style={styles.btnText}>Limpar</Text>
          </TouchableOpacity>
        </View>

        <View style={styles.infoBox}>
          <Text style={[styles.infoText, { color: excelDados.length > 0 ? '#03DAC6' : '#888' }]}>
            📁 Excel: {excelNome} {excelDados.length > 0 ? `(${excelDados.length} nomes)` : ''}
          </Text>
          <Text style={[styles.infoText, { color: imageNome !== "Nenhuma" ? '#BB86FC' : '#888' }]}>
            🖼️ Imagem: {imageNome}
          </Text>
        </View>

        <TextInput
          style={styles.input}
          placeholder="Digite sua mensagem..."
          placeholderTextColor="#555"
          multiline
          value={mensagem}
          onChangeText={setMensagem}
        />

        <Text style={styles.label}>Espera: {tempoCarregamento}s</Text>
        <Slider
          style={styles.slider}
          minimumValue={5} maximumValue={60} step={1}
          value={tempoCarregamento} onValueChange={v => setTempoCarregamento(Math.floor(v))}
          minimumTrackTintColor="#3700B3" thumbTintColor="#BB86FC"
        />

        <Text style={[styles.label, { color: '#03DAC6' }]}>Intervalo: {tempoIntervalo}s</Text>
        <Slider
          style={styles.slider}
          minimumValue={5} maximumValue={120} step={1}
          value={tempoIntervalo} onValueChange={v => setTempoIntervalo(Math.floor(v))}
          minimumTrackTintColor="#03DAC6" thumbTintColor="#03DAC6"
        />

        <TouchableOpacity style={styles.btnGreen} onPress={iniciarEnvio}>
          <Text style={styles.btnTextBold}>▶️ INICIAR ENVIO</Text>
        </TouchableOpacity>

      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#000' },
  scroll: { padding: 20, paddingTop: 60 },
  title: { fontSize: 32, fontWeight: 'bold', color: '#BB86FC', textAlign: 'center', marginBottom: 30 },
  btnBlue: { backgroundColor: '#3700B3', padding: 18, borderRadius: 12, alignItems: 'center', marginBottom: 15 },
  row: { flexDirection: 'row', gap: 10, marginBottom: 12 },
  btnGray: { backgroundColor: '#1A1A1A', padding: 14, borderRadius: 12, alignItems: 'center', borderWidth: 1, borderColor: '#333' },
  btnRed: { backgroundColor: '#CF6679', padding: 14, borderRadius: 12, width: 80, alignItems: 'center' },
  btnGreen: { backgroundColor: '#03DAC6', padding: 22, borderRadius: 15, alignItems: 'center', marginTop: 25 },
  btnText: { color: '#FFF', fontSize: 14, fontWeight: '600' },
  btnTextBold: { color: '#000', fontSize: 16, fontWeight: 'bold' },
  infoBox: { padding: 15, backgroundColor: '#111', borderRadius: 10, marginVertical: 10, borderWidth: 1, borderColor: '#222' },
  infoText: { fontSize: 13, marginBottom: 5, fontWeight: '600' },
  input: { backgroundColor: '#111', borderWidth: 1, borderColor: '#333', borderRadius: 12, padding: 20, height: 150, textAlignVertical: 'top', color: '#FFF', fontSize: 16 },
  label: { fontSize: 13, fontWeight: 'bold', color: '#BB86FC', marginTop: 15 },
  slider: { width: '100%', height: 40 }
});
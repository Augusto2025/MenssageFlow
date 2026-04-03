import React, { useState } from 'react';
import { 
  StyleSheet, Text, View, ScrollView, TouchableOpacity, 
  TextInput, Linking, Alert 
} from 'react-native';
import { StatusBar } from 'expo-status-bar';
import * as DocumentPicker from 'expo-document-picker';
import * as FileSystem from 'expo-file-system/legacy'; 
import Slider from '@react-native-community/slider';
import XLSX from 'xlsx';

export default function App() {
  const [excelNome, setExcelNome] = useState("Nenhum");
  const [excelDados, setExcelDados] = useState([]);
  const [imageNome, setImageNome] = useState("Nenhuma");
  const [mensagem, setMensagem] = useState("");
  const [tempoCarregamento, setTempoCarregamento] = useState(15);
  const [tempoIntervalo, setTempoIntervalo] = useState(30);
  
  const [enviando, setEnviando] = useState(false);
  const [indiceAtual, setIndiceAtual] = useState(0);

  const formatarNumero = (num) => {
    let limpo = String(num).replace(/\D/g, '');
    if (limpo.length >= 10 && limpo.length <= 11) limpo = '55' + limpo;
    return limpo;
  };

  const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

  const carregarExcel = async () => {
    try {
      const result = await DocumentPicker.getDocumentAsync({
        type: ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel", "text/csv"],
        copyToCacheDirectory: true
      });

      if (!result.canceled && result.assets) {
        const file = result.assets[0];
        setExcelNome(file.name);
        const arquivoBase64 = await FileSystem.readAsStringAsync(file.uri, { encoding: 'base64' });
        const workbook = XLSX.read(arquivoBase64, { type: 'base64' });
        const folha = workbook.Sheets[workbook.SheetNames[0]];
        const dados = XLSX.utils.sheet_to_json(folha);

        setExcelDados(dados);
        setIndiceAtual(0);
        Alert.alert("Sucesso", "Planilha carregada!");
      }
    } catch (e) {
      Alert.alert("Erro", "Erro ao ler Excel.");
    }
  };

  const carregarImagem = async () => {
    try {
      const result = await DocumentPicker.getDocumentAsync({ type: "image/*" });
      if (!result.canceled && result.assets) {
        setImageNome(result.assets[0].name);
      }
    } catch (e) {
      Alert.alert("Erro", "Erro ao carregar imagem.");
    }
  };

  // --- FUNÇÃO DE ENVIO ATUALIZADA COM TIMERS ---
  const processarEnvio = async () => {
    if (excelDados.length === 0) return Alert.alert("Erro", "Carregue o Excel!");
    if (!mensagem) return Alert.alert("Erro", "Digite uma mensagem!");
    
    setEnviando(true);

    for (let i = 0; i < excelDados.length; i++) {
      setIndiceAtual(i + 1);
      const contato = excelDados[i];
      const nomeOriginal = contato.nome || contato.Nome || "Cliente";
      const numeroOriginal = contato.numero || contato.Numero || "";

      if (numeroOriginal) {
        const pNome = String(nomeOriginal).trim().split(' ')[0];
        const nLimpo = formatarNumero(numeroOriginal);
        const msgFinal = `Olá ${pNome}!\n${mensagem}`;
        const url = `whatsapp://send?phone=${nLimpo}&text=${encodeURIComponent(msgFinal)}`;

        // 1. Aguarda o "Tempo de Carregamento" antes de abrir o próximo contato
        await sleep(tempoCarregamento * 1000); 

        // 2. Abre o WhatsApp com a mensagem pronta
        const supported = await Linking.canOpenURL(url);
        if (supported) {
          await Linking.openURL(url);
        } else {
          Alert.alert("Erro", "Não foi possível abrir o WhatsApp");
          break;
        }

        // 3. Aguarda o "Tempo de Intervalo" antes de prosseguir no loop
        // Esse é o tempo para você clicar em enviar e voltar para o app manualmente
        await sleep(tempoIntervalo * 1000);
      }
    }

    setEnviando(false);
    setIndiceAtual(0);
    Alert.alert("Finalizado", "Processo concluído!");
  };

  return (
    <View style={styles.container}>
      <StatusBar style="light" />
      <ScrollView contentContainerStyle={styles.scroll}>
        
        <Text style={styles.title}>MessageFlow</Text>

        <TouchableOpacity style={styles.btnBlue} onPress={() => Linking.openURL('whatsapp://')}>
          <Text style={styles.btnTextBold}>1. Abrir WhatsApp</Text>
        </TouchableOpacity>

        {/* Bloco de Excel */}
        <View style={styles.row}>
          <TouchableOpacity style={[styles.btnGray, { flex: 1 }]} onPress={carregarExcel}>
            <Text style={styles.btnText}>2. Selecionar Excel</Text>
          </TouchableOpacity>
          <TouchableOpacity style={styles.btnRed} onPress={() => {setExcelNome("Nenhum"); setExcelDados([]);}}>
            <Text style={styles.btnText}>Limpar</Text>
          </TouchableOpacity>
        </View>

        {/* Bloco de Imagem */}
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
          {enviando && (
            <Text style={styles.statusEnviando}>Processando {indiceAtual} de {excelDados.length}...</Text>
          )}
        </View>

        <TextInput
          style={styles.input}
          placeholder="Mensagem..."
          placeholderTextColor="#444"
          multiline
          value={mensagem}
          onChangeText={setMensagem}
        />

        <Text style={styles.label}>Espera de carregamento: {tempoCarregamento}s</Text>
        <Slider
          style={styles.slider}
          minimumValue={1} maximumValue={60} step={1}
          value={tempoCarregamento} onValueChange={v => setTempoCarregamento(Math.floor(v))}
          minimumTrackTintColor="#BB86FC" thumbTintColor="#BB86FC"
        />

        <Text style={[styles.label, { color: '#03DAC6' }]}>Intervalo entre contatos: {tempoIntervalo}s</Text>
        <Slider
          style={styles.slider}
          minimumValue={5} maximumValue={120} step={1}
          value={tempoIntervalo} onValueChange={v => setTempoIntervalo(Math.floor(v))}
          minimumTrackTintColor="#03DAC6" thumbTintColor="#03DAC6"
        />

        <TouchableOpacity 
          style={[styles.btnGreen, enviando && { opacity: 0.5 }]} 
          onPress={processarEnvio}
          disabled={enviando}
        >
          <Text style={styles.btnTextBold}>
            {enviando ? "PROCESSANDO..." : "▶️ INICIAR ENVIO"}
          </Text>
        </TouchableOpacity>

      </ScrollView>
    </View>
  );
}

const styles = StyleSheet.create({
  container: { flex: 1, backgroundColor: '#000' },
  scroll: { padding: 20, paddingTop: 60 },
  title: { fontSize: 28, fontWeight: 'bold', color: '#BB86FC', textAlign: 'center', marginBottom: 25 },
  btnBlue: { backgroundColor: '#3700B3', padding: 18, borderRadius: 12, alignItems: 'center', marginBottom: 12 },
  row: { flexDirection: 'row', gap: 10, marginBottom: 12 },
  btnGray: { backgroundColor: '#1A1A1A', padding: 14, borderRadius: 12, alignItems: 'center', borderWidth: 1, borderColor: '#333' },
  btnRed: { backgroundColor: '#CF6679', padding: 14, borderRadius: 12, width: 80, alignItems: 'center' },
  btnGreen: { backgroundColor: '#03DAC6', padding: 20, borderRadius: 15, alignItems: 'center', marginTop: 20 },
  btnText: { color: '#FFF', fontSize: 14, fontWeight: '600' },
  btnTextBold: { color: '#000', fontSize: 16, fontWeight: 'bold' },
  infoBox: { padding: 15, backgroundColor: '#0A0A0A', borderRadius: 10, marginVertical: 10, borderWidth: 1, borderColor: '#222' },
  infoText: { fontSize: 13, marginBottom: 5, color: '#FFF' },
  statusEnviando: { color: '#BB86FC', fontWeight: 'bold', marginTop: 10 },
  input: { backgroundColor: '#0A0A0A', borderWidth: 1, borderColor: '#333', borderRadius: 12, padding: 15, height: 120, textAlignVertical: 'top', color: '#FFF', fontSize: 16 },
  label: { fontSize: 12, fontWeight: 'bold', color: '#BB86FC', marginTop: 15 },
  slider: { width: '100%', height: 40 }
});